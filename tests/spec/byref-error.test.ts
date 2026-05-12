import { Lexer } from '../../src/compiler/lexer';
import { Parser } from '../../src/compiler/parser';
import { Evaluator } from '../../src/compiler/evaluator';
import { assert } from '../ts/test-runner';

function evalVBA(code: string): Evaluator {
    const tokens = new Lexer(code).tokenize();
    const ast = new Parser(tokens).parse();
    const ev = new Evaluator(console.log);
    ev.evaluate(ast);
    return ev;
}

const code = `
Sub TestByRefError(ByRef x)
    x = x * 2
    Err.Raise 5 ' Some error
    x = x * 2 ' Should not be reached
End Sub
`;

{
    const ev = evalVBA(code);
    try {
        ev.evalExpression(`
            Dim x
            x = 10
            Call TestByRefError(x)
        `);
    } catch (e) {
        // Expected error
    }
    
    // x should be updated to 20, because the error occurred after x was modified
    // and the finally block synchronizes ByRef variables.
    assert.strictEqual(ev.evalExpression('x'), 20);
    console.log("✅ ByRef synchronizes on unhandled error");
}
