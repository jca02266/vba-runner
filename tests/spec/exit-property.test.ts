import { Lexer } from '../../src/engine/lexer';
import { Parser } from '../../src/engine/parser';
import { Evaluator } from '../../src/engine/evaluator';
import { assert } from '../../test-libs/test-runner';

function runFunc(code: string, name: string, args: any[] = []): any {
    const tokens = new Lexer(code).tokenize();
    const ast = new Parser(tokens).parse();
    const ev = new Evaluator(() => {});
    ev.evaluate(ast);
    return ev.callProcedure(name, args);
}

// 1. Exit Property Statement
{
    const code = `
        Property Get TestProp()
            TestProp = 1
            Exit Property
            TestProp = 2
        End Property
    `;
    // Note: callProcedure might need to handle property names if they are stored differently,
    // but in my implementation it just finds the procedure by name.
    assert.strictEqual(runFunc(code, 'TestProp'), 1, 'Exit Property should terminate execution immediately');
}

console.log('\n✅ Exit Property Statement: 全テスト通過');
