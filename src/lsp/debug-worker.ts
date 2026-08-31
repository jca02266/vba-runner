import { parentPort, workerData, receiveMessageOnPort, MessagePort } from 'worker_threads';
import { Evaluator, Environment, DebugHook } from '../engine/evaluator';
import { vbaToDisplayString } from '../engine/coerce';
import { ProcedureDeclaration } from '../engine/parser';
import { MemoryFileSystem } from '../engine/filesystem';
import { findClassProcedure, findFirstClassProcedure } from '../engine/property-resolution';
import { parseVBAModule } from './vba-source-parser';
import { loadMocks } from '../../test-libs/mock-loader';
import { injectExcelStub } from '../../test-libs/excel-stub';

interface WorkerInitData {
    source: string;
    moduleName: string;
    entryPoint: string | null;
    parseAsClass?: boolean;
    sourceDirectory?: string | null;
    controlBuffer: SharedArrayBuffer;
}

// Commands (control[0]) — must match debug-session.ts
const CMD_WAIT = 0;
const CMD_CONTINUE = 1;
const CMD_STEP_OVER = 2;
const CMD_STEP_INTO = 3;
const CMD_STEP_OUT = 4;
const CMD_TERMINATE = 5;

const { source, moduleName, entryPoint, parseAsClass, sourceDirectory, controlBuffer } = workerData as WorkerInitData;
const control = new Int32Array(controlBuffer);

let currentCommand = CMD_STEP_INTO; // pause at first statement by default
let startCallDepth = 0;
let pauseAfterCurrent = false;
let skipStepLine: number | null = null;
let skipStepDepth: number | null = null;
const breakpointLines = new Set<number>();
const breakpointConditions = new Map<number, string>();
const breakpointHitConditions = new Map<number, number>();
const breakpointHitCounts = new Map<number, number>();
let isFirstPause = true;
let activeEvaluator: Evaluator | null = null;
let isEvaluatingExpression = false;
let lastLine = 0;
let lastFrames: Array<{ id: number; name: string; source: string; line: number; column: number }> = [];
let lastVariables: Array<{ name: string; value: string; type: string }> = [];
let lastFrameVariables: Array<Array<{ name: string; value: string; type: string; variablesReference: number }>> = [];
let nextVariableHandle = 2;
const variableHandles = new Map<number, any>();

function frameEnvironment(evaluator: Evaluator, frameId: number, fallback: Environment): Environment {
    const frames = evaluator.getDebugFrameEnvironments();
    const index = frames.length - 1 - Math.max(0, frameId);
    return frames[index] ?? fallback;
}

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
            breakpointConditions.clear();
            breakpointHitConditions.clear();
            breakpointHitCounts.clear();
            for (const line of msg.lines as number[]) breakpointLines.add(line);
            for (const [line, condition] of Object.entries(msg.conditions ?? {})) {
                if (typeof condition === 'string' && condition.trim()) {
                    breakpointConditions.set(Number(line), condition);
                }
            }
            for (const [line, hitCondition] of Object.entries(msg.hitConditions ?? {})) {
                const threshold = Number.parseInt(String(hitCondition), 10);
                if (Number.isFinite(threshold) && threshold > 0) {
                    breakpointHitConditions.set(Number(line), threshold);
                }
            }
        } else if (msg.type === 'pause') {
            pauseAfterCurrent = true;
        } else if (msg.type === 'evaluate' && env && evaluator) {
            const previousEnv = evaluator.env;
            try {
                evaluator.env = frameEnvironment(evaluator, Number(msg.frameId ?? 0), env);
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
                    variablesReference: variableReference(value),
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
        } else if (msg.type === 'variables') {
            const value = variableHandles.get(Number(msg.variablesReference));
            parentPort!.postMessage({ type: 'variablesResult', requestId: msg.requestId, ok: true, variables: childVariables(value) });
        } else if (msg.type === 'setVariable' && env && evaluator) {
            const previousEnv = evaluator.env;
            try {
                const targetEnv = frameEnvironment(evaluator, Number(msg.frameId ?? 0), env);
                evaluator.env = targetEnv;
                isEvaluatingExpression = true;
                const value = evaluator.evalExpression(`(${String(msg.value ?? '')})`);
                evaluator.env = previousEnv;
                isEvaluatingExpression = false;
                if (!targetEnv.hasVariable(String(msg.name ?? ''))) {
                    throw new Error(`Variable not found: ${String(msg.name ?? '')}`);
                }
                targetEnv.set(String(msg.name ?? ''), value);
                parentPort!.postMessage({
                    type: 'setVariableResult',
                    requestId: msg.requestId,
                    ok: true,
                    result: formatValue(value),
                    valueType: getTypeName(value),
                });
            } catch (error: any) {
                isEvaluatingExpression = false;
                evaluator.env = previousEnv;
                parentPort!.postMessage({
                    type: 'setVariableResult',
                    requestId: msg.requestId,
                    ok: false,
                    error: error?.vbaBareMessage ?? error?.message ?? String(error),
                });
            }
        }
    }
}

function shouldPause(line: number, callDepth: number, breakpointHit = breakpointLines.has(line)): boolean {
    if (pauseAfterCurrent) {
        pauseAfterCurrent = false;
        return true;
    }
    switch (currentCommand) {
        case CMD_CONTINUE:
            return breakpointHit;
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

function variableReference(value: any): number {
    if (value === null || typeof value !== 'object') return 0;
    const ref = nextVariableHandle++;
    variableHandles.set(ref, value);
    return ref;
}

function makeVariable(name: string, value: any) {
    return { name, value: formatValue(value), type: getTypeName(value), variablesReference: variableReference(value) };
}

function childVariables(value: any): any[] {
    if (Array.isArray(value)) {
        const base = Number((value as any).vbaBase ?? 0);
        const dimensions = (value as any).__vbaDimensions__;
        const lower = Number.isFinite(dimensions?.[0]?.lower) ? dimensions[0].lower : base;
        const upper = Number.isFinite(dimensions?.[0]?.upper) ? dimensions[0].upper : value.length - 1;
        const children: any[] = [];
        for (let index = lower; index <= upper; index++) {
            children.push(makeVariable(String(index), value[index]));
        }
        return children;
    }
    if (!value || typeof value !== 'object') return [];
    // Class instances keep their VBA fields in a private Environment rather
    // than as enumerable JavaScript properties. Expose that environment while
    // hiding evaluator bookkeeping and methods.
    const instanceEnv = (value as any).__instanceEnv__;
    if (instanceEnv && typeof instanceEnv.getLocalVariables === 'function') {
        return [...instanceEnv.getLocalVariables()]
            .filter(([name, member]: [string, any]) => !String(name).startsWith('__') && typeof member !== 'function')
            .map(([name, member]: [string, any]) => makeVariable(String(name), member));
    }
    return Object.keys(value)
        .filter(key => !key.startsWith('__') && typeof value[key] !== 'function')
        .map(key => makeVariable(key, value[key]));
}

function extractVariables(env: Environment): Array<{ name: string; value: string; type: string; variablesReference: number }> {
    const vars: Array<{ name: string; value: string; type: string; variablesReference: number }> = [];
    const localVars = env.getLocalVariables();
    for (const [name, value] of localVars) {
        if (typeof value === 'function') continue;
        vars.push(makeVariable(name, value));
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
        // A step command resumes from the statement that just paused. Do not
        // immediately report that same statement again (notably for a Set
        // assignment whose RHS enters a helper procedure).
        if (skipStepLine === line && skipStepDepth === callDepth) {
            skipStepLine = null;
            skipStepDepth = null;
            lastLine = line;
            return;
        }
        lastLine = line;
        lastVariables = extractVariables(env);
        const frameEnvs = activeEvaluator?.getDebugFrameEnvironments() ?? [env];
        lastFrameVariables = frameEnvs.map(frameEnv => extractVariables(frameEnv)).reverse();
        // DAP expects the innermost (currently executing) frame first.
        // The evaluator keeps the call stack in caller-to-callee order, so
        // reverse it only after resolving each caller's call-site line.
        lastFrames = callStack.map((frame, i) => ({
            id: i,
            name: frame.name,
            source: frame.moduleName,
            // A non-leaf frame stores the line at which its child was called.
            // Point the caller at that call site instead of the initial 0.
            line: i === callStack.length - 1 ? line : (callStack[i + 1]?.line ?? frame.line),
            column: 0,
        })).reverse().map((frame, i) => ({ ...frame, id: i }));
        // ステートメント実行前にキュー内のメッセージ（setBreakpoints など）を処理
        processMessages(env, activeEvaluator ?? undefined);
        const pauseRequested = pauseAfterCurrent;
        let breakpointHit = breakpointLines.has(line);
        if (breakpointHit) {
            const hitCount = (breakpointHitCounts.get(line) ?? 0) + 1;
            breakpointHitCounts.set(line, hitCount);
            const threshold = breakpointHitConditions.get(line);
            breakpointHit = threshold === undefined || hitCount >= threshold;
        }
        const condition = breakpointConditions.get(line);
        if (breakpointHit && condition) {
            try {
                isEvaluatingExpression = true;
                const conditionValue = activeEvaluator?.evalExpression(`(${condition})`);
                const primitive = conditionValue && typeof conditionValue === 'object'
                    && typeof (conditionValue as any).valueOf === 'function'
                    ? (conditionValue as any).valueOf()
                    : conditionValue;
                breakpointHit = Boolean(primitive);
            } catch {
                breakpointHit = false;
            } finally {
                isEvaluatingExpression = false;
            }
        }

        if (!shouldPause(line, callDepth, breakpointHit)) return;

        const variables = lastVariables;
        const frames = lastFrames;

        const reason = isFirstPause
            ? 'entry'
            : (pauseRequested ? 'pause' : (breakpointHit ? 'breakpoint' : 'step'));
        isFirstPause = false;

        parentPort!.postMessage({ type: 'paused', line, callDepth, variables, frameVariables: lastFrameVariables, frames, reason });

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
        if (cmd === CMD_STEP_OVER || cmd === CMD_STEP_OUT) {
            skipStepLine = line;
            skipStepDepth = callDepth;
        } else {
            skipStepLine = null;
            skipStepDepth = null;
        }
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
    const excelApplication = injectExcelStub(evaluator);
    const mockModules = sourceDirectory
        ? loadMocks(sourceDirectory, evaluator, { excel: { Application: excelApplication }, sourceDirectory })
        : [];
    for (const { ast: mockAst, moduleName: mockName } of mockModules) {
        evaluator.setSourceModule(mockName);
        evaluator.evaluateModule(mockAst);
    }
    evaluator.setDebugHook(hook);
    evaluator.setSourceModule(moduleName || 'Module1');
    evaluator.evaluateModule(ast);
    evaluator.resolveIdentifiers([{ ast, moduleName: moduleName || 'Module1' }]);

    let ep = entryPoint;
    let classEntry: ProcedureDeclaration | undefined;
    if (parseAsClass && ast.body.length > 0) {
        const classDecl = ast.body.find((stmt: any) => stmt.type === 'ClassDeclaration') as any;
        classEntry = classDecl
            ? (ep
                ? findClassProcedure(classDecl.procedures, ep, 'method')
                : findFirstClassProcedure(classDecl.procedures, 'method'))
            : undefined;
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
        parentPort!.postMessage({
            type: 'error',
            message: msg,
            line: lastLine,
            frames: lastFrames,
            variables: lastVariables,
            frameVariables: lastFrameVariables,
        });
        parentPort!.postMessage({ type: 'exited', exitCode: 1 });
    }
}
