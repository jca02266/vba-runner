import { Program, Statement, ProcedureDeclaration } from '../engine/parser';

export interface Breakpoint {
    id: string;
    line: number;
    column: number;
    verified: boolean;
    message?: string;
}

export interface StackFrame {
    id: number;
    name: string;
    source: string;
    line: number;
    column: number;
}

export interface Variable {
    name: string;
    value: string;
    type: string;
    variablesReference: number;
}

export type DebuggerState = 'initialized' | 'running' | 'stopped' | 'exited';

export class Debugger {
    private ast: Program;
    private breakpoints: Map<string, Breakpoint> = new Map();
    private breakpointCounter: number = 0;
    private state: DebuggerState = 'initialized';
    private stackFrames: StackFrame[] = [];
    private variables: Map<string, any> = new Map();

    constructor(ast: Program) {
        this.ast = ast;
        this.initializeStackFrames();
    }

    /**
     * Set a breakpoint at the given line and column
     */
    setBreakpoint(line: number, column: number): Breakpoint {
        const id = `bp_${++this.breakpointCounter}`;
        const bp: Breakpoint = {
            id,
            line,
            column,
            verified: false, // Will be verified during execution
        };

        // Try to verify: check if line exists in source
        if (this.verifyBreakpointLocation(line)) {
            bp.verified = true;
        }

        this.breakpoints.set(id, bp);
        return bp;
    }

    /**
     * Remove a breakpoint by ID
     */
    removeBreakpoint(id: string): boolean {
        return this.breakpoints.delete(id);
    }

    /**
     * Get all breakpoints
     */
    getBreakpoints(): Breakpoint[] {
        return Array.from(this.breakpoints.values());
    }

    /**
     * Continue execution (resume from breakpoint)
     */
    continue(): void {
        this.state = 'running';
    }

    /**
     * Step over the current statement
     */
    stepOver(): void {
        this.state = 'running';
        // In a real implementation, would execute one statement
    }

    /**
     * Step into a procedure call
     */
    stepInto(): void {
        this.state = 'running';
        // In a real implementation, would enter called procedure
    }

    /**
     * Step out of current procedure
     */
    stepOut(): void {
        this.state = 'running';
        // In a real implementation, would exit current frame
    }

    /**
     * Pause execution
     */
    pause(): void {
        this.state = 'stopped';
    }

    /**
     * Get current debugger state
     */
    getState(): DebuggerState {
        return this.state;
    }

    /**
     * Get current stack frames
     */
    getStackFrames(): StackFrame[] {
        return this.stackFrames;
    }

    /**
     * Get variables at a given frame
     */
    getVariables(_frameId: number): Variable[] {
        // Return variables for the given frame
        // In a real implementation, would extract from execution scope
        const vars: Variable[] = [];

        // Add variables from the AST procedures
        for (const stmt of this.ast.body) {
            if (stmt.type === 'ProcedureDeclaration') {
                const proc = stmt as ProcedureDeclaration;
                // Add parameters as variables
                for (const param of proc.parameters) {
                    vars.push({
                        name: param.name,
                        value: '(parameter)',
                        type: param.paramType || 'Variant',
                        variablesReference: 0,
                    });
                }
            }
        }

        return vars;
    }

    /**
     * Get value of a variable by name
     */
    getVariableValue(name: string): any {
        return this.variables.get(name.toLowerCase());
    }

    /**
     * Set value of a variable
     */
    setVariableValue(name: string, value: any): void {
        this.variables.set(name.toLowerCase(), value);
    }

    /**
     * Evaluate an expression in the current context
     */
    evaluateExpression(_expression: string): any {
        // In a real implementation, would parse and evaluate the expression
        // For now, return undefined
        return undefined;
    }

    /**
     * Get threads (for multi-threaded debugging)
     */
    getThreads(): Array<{ id: number; name: string }> {
        return [{ id: 1, name: 'Main Thread' }];
    }

    private initializeStackFrames(): void {
        // Initialize with main entry point
        const mainFrame: StackFrame = {
            id: 0,
            name: 'Main',
            source: '(module)',
            line: 0,
            column: 0,
        };

        this.stackFrames = [mainFrame];

        // Add procedures as potential frames
        for (const stmt of this.ast.body) {
            if (stmt.type === 'ProcedureDeclaration') {
                const proc = stmt as ProcedureDeclaration;
                const frame: StackFrame = {
                    id: this.stackFrames.length,
                    name: proc.name.name,
                    source: '(module)',
                    line: proc.loc?.start.line ?? 0 - 1,
                    column: proc.loc?.start.column ?? 0 - 1,
                };
                this.stackFrames.push(frame);
            }
        }
    }

    private verifyBreakpointLocation(line: number): boolean {
        // Check if the line exists in the source
        // Look through statements to see if any span this line
        return this.checkLine(this.ast.body, line);
    }

    private checkLine(stmts: Statement[], line: number): boolean {
        for (const stmt of stmts) {
            if (stmt.loc) {
                const startLine = stmt.loc.start.line - 1;
                const endLine = stmt.loc.end.line - 1;
                if (line >= startLine && line <= endLine) {
                    return true;
                }
            }

            // Check nested statements in procedures/classes
            if (stmt.type === 'ProcedureDeclaration') {
                const proc = stmt as ProcedureDeclaration;
                if (this.checkLine(proc.body, line)) {
                    return true;
                }
            } else if (stmt.type === 'ClassDeclaration') {
                const cls = stmt as any;
                if (this.checkLine(cls.body, line)) {
                    return true;
                }
            }
        }

        return false;
    }
}
