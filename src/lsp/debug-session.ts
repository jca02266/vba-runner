import { Worker } from 'worker_threads';
import * as path from 'path';
import * as fs from 'fs';
import { EventEmitter } from 'events';

export interface SessionBreakpoint {
    id: string;
    line: number;
    column: number;
    verified: boolean;
}

export interface SessionStackFrame {
    id: number;
    name: string;
    source: string;
    line: number;
    column: number;
}

export interface SessionVariable {
    name: string;
    value: string;
    type: string;
    variablesReference: number;
}

export type SessionState = 'initialized' | 'running' | 'paused' | 'exited';

// Commands (control[0]) — must match debug-worker.ts
const CMD_WAIT = 0;
const CMD_CONTINUE = 1;
const CMD_STEP_OVER = 2;
const CMD_STEP_INTO = 3;
const CMD_STEP_OUT = 4;
const CMD_TERMINATE = 5;

export class VBADebugSession extends EventEmitter {
    private worker: Worker | null = null;
    private controlBuffer: SharedArrayBuffer;
    private control: Int32Array;
    private state: SessionState = 'initialized';
    private currentLine = 0;
    private currentVariables: SessionVariable[] = [];
    private currentFrames: SessionStackFrame[] = [];
    private breakpoints: Map<string, SessionBreakpoint> = new Map();
    private bpCounter = 0;
    private evaluationCounter = 0;
    private pendingEvaluations = new Map<number, {
        resolve: (value: { result: string; type: string }) => void;
        reject: (reason: Error) => void;
    }>();

    constructor(
        private readonly source: string,
        private readonly moduleName: string,
        private readonly entryPoint: string | null = null,
        private readonly parseAsClass = false
    ) {
        super();
        this.controlBuffer = new SharedArrayBuffer(4);
        this.control = new Int32Array(this.controlBuffer);
        Atomics.store(this.control, 0, CMD_WAIT);
    }

    start(): void {
        const cjsPath = path.join(import.meta.dirname, 'debug-worker.cjs');
        const useCjs = fs.existsSync(cjsPath);

        // production: バンドル済み .cjs をファイルパスで起動
        // development(tsx): tsx CJS フックを eval ワーカー経由でロード
        let worker: Worker;
        const workerData = {
            source: this.source,
            moduleName: this.moduleName,
            entryPoint: this.entryPoint,
            parseAsClass: this.parseAsClass,
            controlBuffer: this.controlBuffer,
        };

        if (useCjs) {
            worker = new Worker(cjsPath, { workerData });
        } else {
            const tsWorkerPath = path.join(import.meta.dirname, 'debug-worker.ts');
            const tsxCjsPath = path.join(import.meta.dirname, '../../node_modules/tsx/dist/cjs/index.cjs');
            // Force the eval worker to CommonJS: package-level ESM otherwise
            // leaves the tsx CJS hook without `require`.
            const evalCode = `require(${JSON.stringify(tsxCjsPath)});require(${JSON.stringify(tsWorkerPath)});`;
            worker = new Worker(evalCode, {
                eval: true,
                workerData,
                execArgv: ['--input-type=commonjs'],
            });
        }
        this.worker = worker;

        // 起動前に設定されたブレークポイントを Worker へ送信
        if (this.breakpoints.size > 0) {
            const lines = [...this.breakpoints.values()].map(bp => bp.line);
            this.worker.postMessage({ type: 'setBreakpoints', lines });
        }

        this.worker.on('message', (msg: any) => {
            switch (msg.type) {
                case 'paused':
                    this.state = 'paused';
                    this.currentLine = msg.line;
                    this.currentVariables = (msg.variables as any[]).map(v => ({
                        ...v,
                        variablesReference: 0,
                    }));
                    this.currentFrames = msg.frames ?? [];
                    this.emit('stopped', { reason: msg.reason ?? 'step', line: msg.line });
                    break;
                case 'output':
                    this.emit('output', msg.text);
                    break;
                case 'exited':
                    this.state = 'exited';
                    this.worker = null;
                    this.emit('exited', msg.exitCode ?? 0);
                    break;
                case 'error':
                    if (msg.line) {
                        this.state = 'paused';
                        this.currentLine = msg.line;
                        this.currentFrames = msg.frames ?? [];
                        this.currentVariables = (msg.variables ?? []).map((v: any) => ({
                            ...v, variablesReference: 0,
                        }));
                        this.emit('stopped', { reason: 'exception', line: msg.line });
                    }
                    this.emit('runtimeError', msg.message);
                    break;
                case 'evaluateResult': {
                    const pending = this.pendingEvaluations.get(msg.requestId);
                    if (!pending) break;
                    this.pendingEvaluations.delete(msg.requestId);
                    if (msg.ok) pending.resolve({ result: msg.result, type: msg.valueType });
                    else pending.reject(new Error(msg.error ?? 'Expression evaluation failed'));
                    break;
                }
            }
        });

        this.worker.on('error', (err: Error) => {
            this.state = 'exited';
            this.worker = null;
            this.emit('runtimeError', err.message);
            this.emit('exited', 1);
        });

        this.state = 'running';
    }

    private sendCommand(cmd: number): void {
        if (!this.worker || this.state !== 'paused') return;
        this.state = 'running';
        Atomics.store(this.control, 0, cmd);
        Atomics.notify(this.control, 0);
    }

    continue(): void { this.sendCommand(CMD_CONTINUE); }
    stepOver(): void { this.sendCommand(CMD_STEP_OVER); }
    stepInto(): void { this.sendCommand(CMD_STEP_INTO); }
    stepOut(): void { this.sendCommand(CMD_STEP_OUT); }

    pause(): void {
        if (this.worker && this.state === 'running') {
            this.worker.postMessage({ type: 'pause' });
        }
    }

    terminate(): void {
        if (this.worker) {
            if (this.state === 'paused') {
                this.sendCommand(CMD_TERMINATE);
            } else {
                this.worker.terminate();
                this.worker = null;
            }
        }
        this.state = 'exited';
    }

    setBreakpoints(lines: number[]): SessionBreakpoint[] {
        this.breakpoints.clear();
        const result: SessionBreakpoint[] = [];
        for (const line of lines) {
            const id = `bp_${++this.bpCounter}`;
            const bp: SessionBreakpoint = { id, line, column: 0, verified: true };
            this.breakpoints.set(id, bp);
            result.push(bp);
        }
        if (this.worker) {
            this.worker.postMessage({ type: 'setBreakpoints', lines });
        }
        return result;
    }

    getState(): SessionState { return this.state; }
    getCurrentLine(): number { return this.currentLine; }
    getVariables(_frameId: number): SessionVariable[] { return this.currentVariables; }
    evaluateExpression(expression: string): Promise<{ result: string; type: string }> {
        if (!this.worker || this.state !== 'paused') {
            return Promise.reject(new Error('Expressions can only be evaluated while paused'));
        }
        const requestId = ++this.evaluationCounter;
        return new Promise((resolve, reject) => {
            this.pendingEvaluations.set(requestId, { resolve, reject });
            this.worker!.postMessage({ type: 'evaluate', requestId, expression });
        });
    }
    getStackFrames(): SessionStackFrame[] { return this.currentFrames; }
    getThreads(): Array<{ id: number; name: string }> {
        return [{ id: 1, name: 'Main Thread' }];
    }
}
