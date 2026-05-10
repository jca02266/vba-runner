import * as fs from 'fs';
import * as path from 'path';
import { Lexer } from '../../src/compiler/lexer';
import { Parser } from '../../src/compiler/parser';
import { Evaluator } from '../../src/compiler/evaluator';

const VBA_EXTENSIONS = new Set(['.vba', '.cls', '.frm']);

export class VBATest {
    private evaluator: Evaluator;

    constructor(pathOrDir: string) {
        this.evaluator = new Evaluator(console.log);

        const stat = fs.statSync(pathOrDir);
        const files = stat.isDirectory()
            ? fs.readdirSync(pathOrDir)
                  .filter(f => VBA_EXTENSIONS.has(path.extname(f).toLowerCase()))
                  .sort()
                  .map(f => path.join(pathOrDir, f))
            : [pathOrDir];

        for (const file of files) {
            const source = fs.readFileSync(file, 'utf-8');
            try {
                const ast = new Parser(new Lexer(source).tokenize()).parse();
                this.evaluator.evaluate(ast);
            } catch (e: any) {
                throw new Error(`[${path.basename(file)}] ${e.message}`);
            }
        }
    }

    run(procedureName: string, args: any[]): any {
        const start = Date.now();
        const result = this.evaluator.callProcedure(procedureName, args);
        const duration = Date.now() - start;
        const formatArgs = args.map(a => typeof a === 'object' ? '[Object]' : String(a)).join(', ');
        console.log(`[PASS] ${procedureName}(${formatArgs}) -> ${result} (${duration}ms)`);
        return result;
    }

    eval(exprString: string): any {
        return this.evaluator.evalExpression(exprString);
    }

    set(name: string, value: any): void {
        this.evaluator.set(name, value);
    }
}

// Keep backward compatibility
export function runVBATest(filePath: string, procedureName: string, args: any[]): any {
    const vbaTest = new VBATest(filePath);
    return vbaTest.run(procedureName, args);
}

// Minimal assert framework
export const assert = {
    strictEqual: (actual: any, expected: any, message?: string) => {
        if (actual !== expected) {
            console.error(`[FAIL] ${message || 'Assertion failed'} - Expected ${expected} but got ${actual}`);
            throw new Error(`Assertion Failed`);
        }
    }
};
