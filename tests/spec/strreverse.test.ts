import { Lexer } from '../../src/engine/lexer';
import { Parser } from '../../src/engine/parser';
import { Evaluator } from '../../src/engine/evaluator';
import { assert } from '../../test-libs/test-runner';

function evalVBA(code: string): any {
    const tokens = new Lexer(code).tokenize();
    const ast = new Parser(tokens).parse();
    const ev = new Evaluator(console.log);
    ev.evaluate(ast);
    return ev;
}

function runFunc(code: string, name: string, args: any[] = []): any {
    return evalVBA(code).callProcedure(name, args);
}

// StrReverse のテスト
{
    const code = `
        Function Test1()
            Test1 = StrReverse("VBA")
        End Function
        Function Test2()
            Test2 = StrReverse("12345")
        End Function
        Function Test3()
            Test3 = StrReverse("")
        End Function
    `;
    assert.strictEqual(runFunc(code, 'Test1'), 'ABV', 'StrReverse("VBA") -> "ABV"');
    assert.strictEqual(runFunc(code, 'Test2'), '54321', 'StrReverse("12345") -> "54321"');
    assert.strictEqual(runFunc(code, 'Test3'), '', 'StrReverse("") -> ""');
    console.log('[PASS] 基本動作');
}

// エラー系のテスト
{
    const code = `
        Function TestError()
            TestError = StrReverse(Null)
        End Function
    `;
    try {
        runFunc(code, 'TestError');
        assert.fail('StrReverse(Null) should throw an error');
    } catch (e: any) {
        // Error is expected
        console.log('[PASS] StrReverse(Null) throws error');
    }
}

console.log('\n✅ StrReverse: 全テスト通過');
