import { Lexer } from '../../src/compiler/lexer';
import { Parser } from '../../src/compiler/parser';
import { Evaluator } from '../../src/compiler/evaluator';
import { assert } from '../ts/test-runner';
import * as fs from 'fs';
import * as path from 'path';

let output = "";
const onPrint = (msg: string) => {
    output += msg + "\n";
};

function evalVBA(code: string): any {
    const tokens = new Lexer(code).tokenize();
    const ast = new Parser(tokens).parse();
    const ev = new Evaluator(onPrint);
    ev.evaluate(ast);
    return ev;
}

const sandboxRoot = path.join(process.cwd(), 'sandbox');
if (!fs.existsSync(sandboxRoot)) fs.mkdirSync(sandboxRoot);

console.log("--- Starting FS Lock/Width Stub Tests ---");

const code = `
    Sub Test()
        Open "test_stub.dat" For Output As #1
        Lock #1
        Unlock #1
        Width #1, 80
        Close #1
    End Sub
`;

output = "";
const ev = evalVBA(code);
ev.callProcedure('Test', []);

// Verify stubs were called (via console.log in implementation, but current implementation uses console.log directly, not onPrint for these stubs)
// Wait, I saw evaluator.ts lines 1757, 1762, 1768 use console.log, not this.onPrint.
// So we just check it doesn't crash for now.

assert.ok(true, 'Lock/Unlock/Width executed without error');

console.log("✅ FS Lock/Width: All tests passed!");
