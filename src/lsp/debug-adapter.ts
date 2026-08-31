import { VBADebugSession, SessionBreakpoint, SessionStackFrame, SessionVariable } from './debug-session';

export interface DebugCapabilities {
    supportsConfigurationDoneRequest?: boolean;
    supportsSetVariable?: boolean;
    supportsEvaluateForHovers?: boolean;
    supportsStepBack?: boolean;
    supportsGotoTargetsRequest?: boolean;
    supportsStepInTargetsRequest?: boolean;
    supportsCompletionsRequest?: boolean;
    supportTerminateDebuggee?: boolean;
    supportsDelayedStackTraceLoading?: boolean;
    supportsLoadedSourcesRequest?: boolean;
}

export interface DebugRequest {
    command: string;
    arguments?: any;
}

export interface DebugResponse {
    success: boolean;
    body?: any;
    error?: string;
}

export type { SessionBreakpoint as Breakpoint, SessionStackFrame as StackFrame, SessionVariable as Variable };

export class DebugAdapter {
    private session: VBADebugSession | null = null;
    private initialized = false;
    private moduleName: string;
    private pendingBreakpointLines: number[] = [];
    private pendingBreakpointConditions = new Map<number, string>();
    private pendingBreakpointHitConditions = new Map<number, string>();

    /** DAP イベントを送るコールバック（VBAInlineDebugAdapter が設定する） */
    public onEvent?: (event: any) => void;

    constructor(
        private readonly source: string,
        moduleName = 'Module1',
        private readonly filePath: string | null = null,
        private readonly parseAsClass = false,
        private readonly sourceDirectory: string | null = null
    ) {
        this.moduleName = moduleName;
    }

    handleInitialize(): any {
        this.initialized = true;
        return {
            capabilities: {
                supportsConfigurationDoneRequest: true,
                supportsSetVariable: true,
                supportsEvaluateForHovers: true,
                supportsStepBack: false,
                supportsGotoTargetsRequest: false,
                supportsStepInTargetsRequest: true,
                supportsCompletionsRequest: false,
            } as DebugCapabilities,
        };
    }

    handleLaunch(args: any): any {
        if (!this.initialized) {
            return { success: false, error: 'Not initialized' };
        }

        const entryPoint = args?.entryPoint ?? null;
        this.session = new VBADebugSession(this.source, this.moduleName, entryPoint, this.parseAsClass, this.sourceDirectory);
        if (this.pendingBreakpointLines.length > 0) {
            this.session.setBreakpoints(
                this.pendingBreakpointLines,
                this.pendingBreakpointConditions,
                this.pendingBreakpointHitConditions,
            );
        }

        this.session.on('stopped', (info: { reason: string; line: number }) => {
            this.onEvent?.({
                type: 'event',
                event: 'stopped',
                body: { reason: info.reason, threadId: 1, allThreadsStopped: true },
            });
        });

        this.session.on('output', (text: string) => {
            this.onEvent?.({
                type: 'event',
                event: 'output',
                body: { category: 'stdout', output: text + '\n' },
            });
        });

        this.session.on('runtimeError', (msg: string) => {
            this.onEvent?.({
                type: 'event',
                event: 'output',
                body: { category: 'stderr', output: msg + '\n' },
            });
        });

        this.session.on('exited', (exitCode: number) => {
            this.onEvent?.({ type: 'event', event: 'terminated' });
            this.onEvent?.({ type: 'event', event: 'exited', body: { exitCode } });
        });

        this.session.start();
        return { success: true };
    }

    handleConfigurationDone(): any {
        return { success: true };
    }

    handleSetBreakpoints(args: any): any {
        const lines: number[] = [];
        const conditions = new Map<number, string>();
        const hitConditions = new Map<number, string>();
        if (Array.isArray(args?.breakpoints)) {
            for (const bp of args.breakpoints) {
                lines.push(bp.line);
                if (typeof bp.condition === 'string' && bp.condition.trim()) {
                    conditions.set(bp.line, bp.condition);
                }
                if (typeof bp.hitCondition === 'string' && bp.hitCondition.trim()) {
                    hitConditions.set(bp.line, bp.hitCondition);
                }
            }
        } else if (args?.line !== undefined) {
            lines.push(args.line);
        }

        this.pendingBreakpointLines = [...lines];
        this.pendingBreakpointConditions = conditions;
        this.pendingBreakpointHitConditions = hitConditions;

        const bps = this.session?.setBreakpoints(lines, conditions, hitConditions) ?? lines.map((line, i) => ({
            id: `bp_${i}`,
            line,
            column: 0,
            verified: false,
            condition: conditions.get(line),
            hitCondition: hitConditions.get(line),
        }));

        return {
            breakpoints: bps.map(bp => ({
                id: bp.id, verified: bp.verified, line: bp.line,
                ...(bp.condition ? { condition: bp.condition } : {}),
                ...(bp.hitCondition ? { hitCondition: bp.hitCondition } : {}),
            })),
        };
    }

    handleThreads(): Array<{ id: number; name: string }> {
        return this.session?.getThreads() ?? [{ id: 1, name: 'Main Thread' }];
    }

    handleStackTrace(_threadId: number): any {
        const frames = this.session?.getStackFrames() ?? [];
        const stackFrames = frames.map(frame => {
            const source = this.filePath && frame.source === this.moduleName
                ? { name: this.moduleName, path: this.filePath }
                : { name: frame.source };
            return { ...frame, source };
        });
        return { stackFrames, totalFrames: stackFrames.length };
    }

    handleVariables(variablesReference: number): any {
        if (variablesReference === 1) {
            return { variables: this.session?.getVariables(0) ?? [] };
        }
        if (!this.session) return { variables: [] };
        return this.session.requestVariables(variablesReference).then(variables => ({ variables }));
    }

    handleScopes(_frameId: number): any {
        return {
            scopes: [{ name: 'Locals', variablesReference: 1, expensive: false }],
        };
    }

    handleContinue(_threadId: number): any {
        this.session?.continue();
        return { allThreadsContinued: true };
    }

    handleStepOver(_threadId: number): any {
        this.session?.stepOver();
        return {};
    }

    handleStepInto(_threadId: number): any {
        this.session?.stepInto();
        return {};
    }

    handleStepOut(_threadId: number): any {
        this.session?.stepOut();
        return {};
    }

    handlePause(_threadId: number): any {
        this.session?.pause();
        return {};
    }

    async handleEvaluate(expression: string, _frameId: number, _context?: string): Promise<any> {
        if (!this.session) {
            return { success: false, error: 'Debug session is not running' };
        }
        try {
            const value = await this.session.evaluateExpression(expression);
            return { ...value, variablesReference: value.variablesReference ?? 0 };
        } catch (error: any) {
            return { success: false, error: error?.message ?? String(error) };
        }
    }

    async handleSetVariable(_frameId: number, name: string, value: string): Promise<any> {
        if (!this.session) {
            return { success: false, error: 'Debug session is not running' };
        }
        try {
            const result = await this.session.setVariable(name, value);
            return { ...result, variablesReference: 0 };
        } catch (error: any) {
            return { success: false, error: error?.message ?? String(error) };
        }
    }

    handleDisconnect(): any {
        this.session?.terminate();
        this.session = null;
        return { success: true };
    }

    async handleRequest(request: DebugRequest): Promise<any> {
        switch (request.command) {
            case 'initialize':
                return this.handleInitialize();
            case 'launch':
                return this.handleLaunch(request.arguments);
            case 'configurationDone':
                return this.handleConfigurationDone();
            case 'setBreakpoints':
                return this.handleSetBreakpoints(request.arguments);
            case 'threads':
                return { threads: this.handleThreads() };
            case 'stackTrace':
                return this.handleStackTrace(request.arguments?.threadId ?? 1);
            case 'variables':
                return await this.handleVariables(request.arguments?.variablesReference ?? 0);
            case 'scopes':
                return this.handleScopes(request.arguments?.frameId ?? 0);
            case 'continue':
                return this.handleContinue(request.arguments?.threadId ?? 1);
            case 'next':
                return this.handleStepOver(request.arguments?.threadId ?? 1);
            case 'stepIn':
                return this.handleStepInto(request.arguments?.threadId ?? 1);
            case 'stepOut':
                return this.handleStepOut(request.arguments?.threadId ?? 1);
            case 'pause':
                return this.handlePause(request.arguments?.threadId ?? 1);
            case 'evaluate':
                return this.handleEvaluate(
                    request.arguments?.expression ?? '',
                    request.arguments?.frameId ?? 0,
                    request.arguments?.context
                );
            case 'setVariable':
                return this.handleSetVariable(
                    request.arguments?.frameId ?? 0,
                    request.arguments?.name ?? '',
                    request.arguments?.value ?? ''
                );
            case 'disconnect':
                return this.handleDisconnect();
            default:
                return { success: false, error: `Unknown command: ${request.command}` };
        }
    }
}
