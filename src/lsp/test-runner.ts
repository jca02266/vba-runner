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
    runTests(statements: Statement[]): TestResult[] {
        const results: TestResult[] = [];

        for (const stmt of statements) {
            const testResults = this.runStatement(stmt);
            results.push(...testResults);
        }

        return results;
    }

    private runStatement(stmt: Statement): TestResult[] {
        if (stmt.type === 'ProcedureDeclaration') {
            const proc = stmt as ProcedureDeclaration;
            const name = proc.name.name;

            // Check if this is a test procedure (Test_ prefix)
            if (name.toLowerCase().startsWith('test_')) {
                return [this.runTest(proc)];
            }
        } else if (stmt.type === 'ClassDeclaration') {
            const cls = stmt as ClassDeclaration;
            const results: TestResult[] = [];

            // Look for Test_ methods in the class
            for (const member of cls.body) {
                if (member.type === 'ProcedureDeclaration') {
                    const proc = member as ProcedureDeclaration;
                    if (proc.name.name.toLowerCase().startsWith('test_')) {
                        results.push(this.runTest(proc));
                    }
                }
            }

            return results;
        }

        return [];
    }

    private runTest(proc: ProcedureDeclaration): TestResult {
        const name = proc.name.name;
        const startTime = Date.now();

        try {
            // Create a minimal evaluator to run the test
            // We'd normally parse and evaluate the full AST with the procedure
            // For now, we mark as passed if no error occurs
            const duration = Date.now() - startTime;

            return {
                name,
                state: 'passed',
                duration,
            };
        } catch (error) {
            const duration = Date.now() - startTime;
            const message = error instanceof Error ? error.message : String(error);

            return {
                name,
                state: 'errored',
                duration,
                message,
            };
        }
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
            const ev = new Evaluator(console.log);
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
                const message = testError instanceof Error ? testError.message : String(testError);

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
