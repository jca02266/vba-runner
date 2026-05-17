import { Lexer } from '../../src/engine/lexer';
import { Parser } from '../../src/engine/parser';
import { Evaluator } from '../../src/engine/evaluator';
import { assert } from '../../test-libs/test-runner';

function runFunc(code: string, name: string, args: any[] = []): any {
    const tokens = new Lexer(code).tokenize();
    const ast = new Parser(tokens).parse();
    const ev = new Evaluator(() => {});
    ev.evaluate(ast);
    try {
        return ev.callProcedure(name, args);
    } catch (e: any) {
        if (e && e.type === 'Terminate') {
            return 'TERMINATED';
        }
        throw e;
    }
}

// 1. Stop Statement (基本的にはログ出力して継続)
{
    const code = `
        Function TestStop()
            Stop
            TestStop = 123
        End Function
    `;
    assert.strictEqual(runFunc(code, 'TestStop'), 123, 'Stop should not crash the program');
}

// 2. End Statement (即時終了)
{
    const code = `
        Function TestEnd()
            TestEnd = 1
            End
            TestEnd = 2
        End Function
    `;
    assert.strictEqual(runFunc(code, 'TestEnd'), 'TERMINATED', 'End should terminate execution immediately');
}

console.log('\n✅ Stop & End Statements: 全テスト通過');
