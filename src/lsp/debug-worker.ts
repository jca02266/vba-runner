import { parentPort, workerData, receiveMessageOnPort, MessagePort } from 'worker_threads';
import { Evaluator, Environment, DebugHook } from '../engine/evaluator';
import { vbaToDisplayString } from '../engine/coerce';
import { ProcedureDeclaration } from '../engine/parser';
import { MemoryFileSystem } from '../engine/filesystem';
import { parseVBAModule } from './vba-source-parser';

interface WorkerInitData {
    source: string;
    moduleName: string;
    entryPoint: string | null;
    parseAsClass?: boolean;
    controlBuffer: SharedArrayBuffer;
}

// Commands (control[0]) — must match debug-session.ts
const CMD_WAIT = 0;
const CMD_CONTINUE = 1;
const CMD_STEP_OVER = 2;
const CMD_STEP_INTO = 3;
const CMD_STEP_OUT = 4;
const CMD_TERMINATE = 5;

const { source, moduleName, entryPoint, parseAsClass, controlBuffer } = workerData as WorkerInitData;
const control = new Int32Array(controlBuffer);

let currentCommand = CMD_STEP_INTO; // pause at first statement by default
let startCallDepth = 0;
let pauseAfterCurrent = false;
const breakpointLines = new Set<number>();
let isFirstPause = true;
let activeEvaluator: Evaluator | null = null;
let isEvaluatingExpression = false;

/**
 * Atomics.wait はスレッドをブロックするためイベントループが止まり、
 * parentPort の 'message' イベントが発火しない。
 * receiveMessageOnPort で同期的にキューを読み、setBreakpoints / pause を処理する。
 */
function processMessages(env?: Environment, evaluator?: Evaluator): void {
    while (true) {
        const received = receiveMessageOnPort(parentPort as MessagePort);
        if (!received) break;
        const msg = received.message;
        if (msg.type === 'setBreakpoints') {
            breakpointLines.clear();
            for (const line of msg.lines as number[]) breakpointLines.add(line);
        } else if (msg.type === 'pause') {
            pauseAfterCurrent = true;
        } else if (msg.type === 'evaluate' && env && evaluator) {
            const previousEnv = evaluator.env;
            try {
                evaluator.env = env;
                isEvaluatingExpression = true;
                // DAP evaluate is explicitly an expression context. Wrapping
                // it removes VBA's statement-vs-expression ambiguity for
                // inputs such as `x + 1`.
                const value = evaluator.evalExpression(`(${String(msg.expression ?? '')})`);
                isEvaluatingExpression = false;
                evaluator.env = previousEnv;
                parentPort!.postMessage({
                    type: 'evaluateResult',
                    requestId: msg.requestId,
                    ok: true,
                    result: formatValue(value),
                    valueType: getTypeName(value),
                });
            } catch (error: any) {
                isEvaluatingExpression = false;
                evaluator.env = previousEnv;
                parentPort!.postMessage({
                    type: 'evaluateResult',
                    requestId: msg.requestId,
                    ok: false,
                    error: error?.vbaBareMessage ?? error?.message ?? String(error),
                });
            }
        }
    }
}

function shouldPause(line: number, callDepth: number): boolean {
    if (pauseAfterCurrent) {
        pauseAfterCurrent = false;
        return true;
    }
    switch (currentCommand) {
        case CMD_CONTINUE:
            return breakpointLines.has(line);
        case CMD_STEP_OVER:
            return callDepth <= startCallDepth;
        case CMD_STEP_INTO:
            return true;
        case CMD_STEP_OUT:
            return callDepth < startCallDepth;
        default:
            return false;
    }
}

function formatValue(value: any): string {
    if (value === null || value === undefined) return '(null)';
    if (typeof value === 'function') return '(function)';
    if (Array.isArray(value)) return `Array(${value.length})`;
    if (typeof value === 'object' && value !== null) {
        if (value.__isVbaBoolean__) return value.valueOf() ? 'True' : 'False';
        return '(Object)';
    }
    try {
        return vbaToDisplayString(value);
    } catch {
        return String(value);
    }
}

function getTypeName(value: any): string {
    if (value === null || value === undefined) return 'Empty';
    if (typeof value === 'function') return 'Function';
    if (Array.isArray(value)) return 'Array';
    if (typeof value === 'object' && value !== null) {
        if (value.__isVbaBoolean__) return 'Boolean';
        return 'Object';
    }
    if (typeof value === 'string') return 'String';
    if (typeof value === 'number') return Number.isInteger(value) ? 'Long' : 'Double';
    if (typeof value === 'boolean') return 'Boolean';
    return 'Variant';
}

function extractVariables(env: Environment): Array<{ name: string; value: string; type: string }> {
    const vars: Array<{ name: string; value: string; type: string }> = [];
    const localVars = env.getLocalVariables();
    for (const [name, value] of localVars) {
        if (typeof value === 'function') continue;
        vars.push({ name, value: formatValue(value), type: getTypeName(value) });
    }
    return vars;
}

const hook: DebugHook = {
    onBeforeStatement(
        line: number,
        callDepth: number,
        env: Environment,
        callStack: ReadonlyArray<{ name: string; moduleName: string; line: number }>
    ) {
        if (isEvaluatingExpression) return;
        // ステートメント実行前にキュー内のメッセージ（setBreakpoints など）を処理
        processMessages(env, activeEvaluator ?? undefined);

        if (!shouldPause(line, callDepth)) return;

        const variables = extractVariables(env);
        const frames = callStack.map((frame, i) => ({
            id: i,
            name: frame.name,
            source: frame.moduleName,
            line: i === callStack.length - 1 ? line : frame.line,
            column: 0,
        }));

        const reason = isFirstPause ? 'entry' : (breakpointLines.has(line) ? 'breakpoint' : 'step');
        isFirstPause = false;

        parentPort!.postMessage({ type: 'paused', line, callDepth, variables, frames, reason });

        // main スレッドからのコマンドを待つ。タイムアウトを設けて
        // evaluate要求などのMessagePortも停止中に処理する。
        while (Atomics.load(control, 0) === CMD_WAIT) {
            processMessages(env, activeEvaluator ?? undefined);
            Atomics.wait(control, 0, CMD_WAIT, 50);
        }
        const cmd = Atomics.load(control, 0);

        if (cmd === CMD_TERMINATE) {
            throw { type: 'DebugTerminate' };
        }

        currentCommand = cmd;
        startCallDepth = callDepth;
        Atomics.store(control, 0, CMD_WAIT);

        // Atomics.wait 復帰後のキューも処理（Resume 直後の setBreakpoints など）
        processMessages(env, activeEvaluator ?? undefined);
    },
};

try {
    const ast = parseVBAModule(source, { moduleName: moduleName || 'Module1', isClass: parseAsClass });
    for (const d of ast.diagnostics) {
        parentPort!.postMessage({ type: 'output', text: `[parse warning] Line ${d.loc.start.line}: ${d.message}` });
    }

    const evaluator = new Evaluator(
        (text) => parentPort!.postMessage({ type: 'output', text }),
        { fs: new MemoryFileSystem(), allowTopLevelStatements: false }
    );
    activeEvaluator = evaluator;
    evaluator.setDebugHook(hook);
    evaluator.setSourceModule(moduleName || 'Module1');
    evaluator.evaluateModule(ast);
    evaluator.resolveIdentifiers([{ ast, moduleName: moduleName || 'Module1' }]);

    let ep = entryPoint;
    let classEntry: ProcedureDeclaration | undefined;
    if (parseAsClass && ast.body.length > 0) {
        const classDecl = ast.body.find((stmt: any) => stmt.type === 'ClassDeclaration') as any;
        classEntry = classDecl?.procedures?.find((proc: ProcedureDeclaration) =>
            !proc.isProperty && (!ep || proc.name.name.toLowerCase() === ep.toLowerCase())
        );
        if (!ep && classEntry) ep = classEntry.name.name;
    }
    if (!ep) {
        // まず Sub を探し、なければ Function にフォールバック
        let firstFunction: string | null = null;
        for (const stmt of ast.body) {
            if (stmt.type === 'ProcedureDeclaration') {
                const proc = stmt as ProcedureDeclaration;
                if (proc.isProperty) continue;
                if (!proc.isFunction) {
                    ep = proc.name.name;
                    break;
                }
                if (!firstFunction) firstFunction = proc.name.name;
            }
        }
        if (!ep) ep = firstFunction;
    }

    if (ep && parseAsClass) {
        const wrapperSource = `Sub __VbaDebugClassEntry()\n    Dim __vbaDebugInstance As ${moduleName}\n    Set __vbaDebugInstance = New ${moduleName}\n    Call __vbaDebugInstance.${ep}()\nEnd Sub`;
        const wrapperAst = parseVBAModule(wrapperSource, { moduleName: '__VbaDebugEntry' });
        evaluator.evaluateModule(wrapperAst);
        evaluator.resolveIdentifiers([
            { ast, moduleName: moduleName || 'Module1' },
            { ast: wrapperAst, moduleName: '__VbaDebugEntry' },
        ]);
        evaluator.callProcedure('__VbaDebugClassEntry', []);
    } else if (ep) {
        evaluator.callProcedure(ep, []);
    } else {
        parentPort!.postMessage({ type: 'output', text: 'Error: No Sub or Function found. Add "entryPoint" to your launch.json configuration (e.g. "entryPoint": "MyProcedure").' });
    }

    parentPort!.postMessage({ type: 'exited', exitCode: 0 });
} catch (e: any) {
    if (e && e.type === 'DebugTerminate') {
        parentPort!.postMessage({ type: 'exited', exitCode: 0 });
    } else {
        const msg = e instanceof Error ? e.message : String(e);
        parentPort!.postMessage({ type: 'error', message: msg });
        parentPort!.postMessage({ type: 'exited', exitCode: 1 });
    }
}
