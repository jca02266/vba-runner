import { Program } from '../engine/parser';
import { Debugger, Breakpoint, StackFrame, Variable } from './debugger';

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

export class DebugAdapter {
    private debugger: Debugger;
    private initialized: boolean = false;

    constructor(ast: Program) {
        this.debugger = new Debugger(ast);
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

    handleLaunch(_args: any): any {
        if (!this.initialized) {
            return { success: false, error: 'Not initialized' };
        }

        // Initialize the debugger session
        this.debugger.getState();
        return { success: true };
    }

    handleSetBreakpoints(location: any): Breakpoint[] {
        const breakpoints: Breakpoint[] = [];

        if (location.line !== undefined) {
            const bp = this.debugger.setBreakpoint(location.line, location.column ?? 0);
            breakpoints.push(bp);
        }

        return breakpoints;
    }

    handleThreads(): Array<{ id: number; name: string }> {
        return this.debugger.getThreads();
    }

    handleStackTrace(_threadId: number): StackFrame[] {
        return this.debugger.getStackFrames();
    }

    handleVariables(frameId: number): Variable[] {
        return this.debugger.getVariables(frameId);
    }

    handleContinue(_threadId: number): boolean {
        this.debugger.continue();
        return true;
    }

    handleStepOver(_threadId: number): any {
        this.debugger.stepOver();
        return { success: true };
    }

    handleStepInto(_threadId: number): any {
        this.debugger.stepInto();
        return { success: true };
    }

    handleStepOut(_threadId: number): any {
        this.debugger.stepOut();
        return { success: true };
    }

    handlePause(_threadId: number): any {
        this.debugger.pause();
        return { success: true };
    }

    handleEvaluate(expression: string, _frameId: number, _context?: string): any {
        const result = this.debugger.evaluateExpression(expression);

        return {
            result: result !== undefined ? String(result) : '(undefined)',
            type: typeof result,
            variablesReference: 0,
        };
    }

    handleSetVariable(_frameId: number, name: string, value: string): any {
        try {
            // Try to parse value as a number, otherwise treat as string
            let parsedValue: any = value;
            if (!isNaN(Number(value))) {
                parsedValue = Number(value);
            }

            this.debugger.setVariableValue(name, parsedValue);

            return {
                value: String(parsedValue),
                type: typeof parsedValue,
                variablesReference: 0,
            };
        } catch (error) {
            return {
                success: false,
                error: error instanceof Error ? error.message : String(error),
            };
        }
    }

    handleDisconnect(): any {
        return { success: true };
    }

    handleRequest(request: DebugRequest): any {
        switch (request.command) {
            case 'initialize':
                return this.handleInitialize();
            case 'launch':
                return this.handleLaunch(request.arguments);
            case 'setBreakpoints':
                return this.handleSetBreakpoints(request.arguments);
            case 'threads':
                return this.handleThreads();
            case 'stackTrace':
                return this.handleStackTrace(request.arguments?.threadId ?? 1);
            case 'variables':
                return this.handleVariables(request.arguments?.frameId ?? 0);
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
