import { Lexer } from '../../src/compiler/lexer';
import { Parser } from '../../src/compiler/parser';
import { Evaluator } from '../../src/compiler/evaluator';
import { assert } from '../ts/test-runner';

function evalVBA(code: string): any {
    const tokens = new Lexer(code).tokenize();
    const ast = new Parser(tokens).parse();
    const ev = new Evaluator(console.log);
    ev.evaluate(ast);
    return ev;
}

console.log("--- Starting Option Private Module Tests ---");

const code = `
    Option Private Module
    
    Public Function Hello()
        Hello = "World"
    End Function
`;

const ev = evalVBA(code);
assert.strictEqual(ev.callProcedure('Hello', []), "World", 'Option Private Module allowed and code executes');

console.log("✅ Option Private Module: All tests passed!");
