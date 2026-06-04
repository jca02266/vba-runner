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

    /** DAP イベントを送るコールバック（VBAInlineDebugAdapter が設定する） */
    public onEvent?: (event: any) => void;

    constructor(private readonly source: string, moduleName = 'Module1') {
        this.moduleName = moduleName;
    }

    handleInitialize(): any {
        this.initialized = true;
        return {
            capabilities: {
                supportsConfigurationDoneRequest: true,
                supportsSetVariable: false,
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
        this.session = new VBADebugSession(this.source, this.moduleName, entryPoint);

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
        if (Array.isArray(args?.breakpoints)) {
            for (const bp of args.breakpoints) lines.push(bp.line);
        } else if (args?.line !== undefined) {
            lines.push(args.line);
        }

        const bps = this.session?.setBreakpoints(lines) ?? lines.map((line, i) => ({
            id: `bp_${i}`,
            line,
            column: 0,
            verified: false,
        }));

        return {
            breakpoints: bps.map(bp => ({ id: bp.id, verified: bp.verified, line: bp.line })),
        };
    }

    handleThreads(): Array<{ id: number; name: string }> {
        return this.session?.getThreads() ?? [{ id: 1, name: 'Main Thread' }];
    }

    handleStackTrace(_threadId: number): any {
        const frames = this.session?.getStackFrames() ?? [];
        return { stackFrames: frames, totalFrames: frames.length };
    }

    handleVariables(_frameId: number): any {
        const vars = this.session?.getVariables(_frameId) ?? [];
        return { variables: vars };
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

    handleEvaluate(_expression: string, _frameId: number, _context?: string): any {
        return {
            result: '(evaluation not supported in current version)',
            type: 'string',
            variablesReference: 0,
        };
    }

    handleSetVariable(_frameId: number, _name: string, _value: string): any {
        return { success: false, error: 'setVariable not supported' };
    }

    handleDisconnect(): any {
        this.session?.terminate();
        this.session = null;
        return { success: true };
    }

    handleRequest(request: DebugRequest): any {
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
                return this.handleVariables(request.arguments?.variablesReference ?? 0);
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
