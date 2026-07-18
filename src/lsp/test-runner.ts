import { Lexer } from '../engine/lexer';
import { Parser } from '../engine/parser';
import { Evaluator } from '../engine/evaluator';
import { ProcedureDeclaration, Statement, ClassDeclaration } from '../engine/parser';

export type TestState = 'passed' | 'failed' | 'errored' | 'skipped';

export interface TestResult {
    name: string;
    state: TestState;
    duration: number;
    message?: string;
}

export class TestRunner {
    constructor(private onPrint: (s: string) => void = console.log) {}

    /**
     * Parse `src`, discover all Test_* procedures, evaluate the module once,
     * then call each test procedure and collect results.
     */
    runTests(src: string): TestResult[] {
        const tokens = new Lexer(src).tokenize();
        const ast = new Parser(tokens).parse();

        const testProcs: ProcedureDeclaration[] = [];
        this.collectTestProcs(ast.body, testProcs);
        if (testProcs.length === 0) return [];

        const ev = new Evaluator(this.onPrint, { allowTopLevelStatements: false });
        ev.evaluateModule(ast);
        ev.resolveIdentifiers([{ ast, moduleName: '' }]);

        const results: TestResult[] = [];
        for (const proc of testProcs) {
            const name = proc.name.name;
            const startTime = Date.now();
            try {
                ev.callProcedure(name, []);
                results.push({ name, state: 'passed', duration: Date.now() - startTime });
            } catch (error) {
                const message = (error as any)?.vbaBareMessage ?? (error as any)?.message ?? String(error);
                results.push({ name, state: 'failed', duration: Date.now() - startTime, message });
            }
        }
        return results;
    }

    /**
     * Run a test procedure with full evaluation context
     * Returns true if test passes, false otherwise
     */
    runTestWithEvaluation(src: string, testName: string): TestResult {
        const startTime = Date.now();

        try {
            const tokens = new Lexer(src).tokenize();
            const ast = new Parser(tokens).parse();

            // Find the test procedure
            const testProc = this.findProcedure(ast.body, testName);
            if (!testProc) {
                const duration = Date.now() - startTime;
                return {
                    name: testName,
                    state: 'failed',
                    duration,
                    message: `Test procedure ${testName} not found`,
                };
            }

            // Evaluate the entire module to set up scope, then call the test
            const ev = new Evaluator(this.onPrint, { allowTopLevelStatements: false });
            ev.evaluateModule(ast);
            ev.resolveIdentifiers([{ ast, moduleName: '' }]);

            // Call the test procedure
            try {
                ev.callProcedure(testName, []);
                const duration = Date.now() - startTime;

                return {
                    name: testName,
                    state: 'passed',
                    duration,
                };
            } catch (testError) {
                const duration = Date.now() - startTime;
                const message = (testError as any)?.vbaBareMessage ?? (testError as any)?.message ?? String(testError);

                return {
                    name: testName,
                    state: 'failed',
                    duration,
                    message,
                };
            }
        } catch (parseError) {
            const duration = Date.now() - startTime;
            const message = parseError instanceof Error ? parseError.message : String(parseError);

            return {
                name: testName,
                state: 'errored',
                duration,
                message,
            };
        }
    }

    private collectTestProcs(statements: Statement[], out: ProcedureDeclaration[]): void {
        for (const stmt of statements) {
            if (stmt.type === 'ProcedureDeclaration') {
                const proc = stmt as ProcedureDeclaration;
                if (proc.name.name.toLowerCase().startsWith('test_')) {
                    out.push(proc);
                }
            } else if (stmt.type === 'ClassDeclaration') {
                const cls = stmt as ClassDeclaration;
                for (const member of cls.body) {
                    if (member.type === 'ProcedureDeclaration') {
                        const proc = member as ProcedureDeclaration;
                        if (proc.name.name.toLowerCase().startsWith('test_')) {
                            out.push(proc);
                        }
                    }
                }
            }
        }
    }

    private findProcedure(statements: Statement[], name: string): ProcedureDeclaration | null {
        for (const stmt of statements) {
            if (stmt.type === 'ProcedureDeclaration') {
                const proc = stmt as ProcedureDeclaration;
                if (proc.name.name.toLowerCase() === name.toLowerCase()) {
                    return proc;
                }
            } else if (stmt.type === 'ClassDeclaration') {
                const cls = stmt as ClassDeclaration;
                for (const member of cls.body) {
                    if (member.type === 'ProcedureDeclaration') {
                        const proc = member as ProcedureDeclaration;
                        if (proc.name.name.toLowerCase() === name.toLowerCase()) {
                            return proc;
                        }
                    }
                }
            }
        }

        return null;
    }
}
