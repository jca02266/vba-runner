import { evalVBASingle, assert } from '../../test-libs/test-runner';

function evalVBA(code: string): any {
    return evalVBASingle(code);
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

// Bug CZ: StrReverse(Null) → エラー (§6.1.2.11.1.40)
{
    const errOf = (expr: string) => {
        try { evalVBASingle('').evalExpression(expr); return null; } catch(e: any) { return e.number ?? -1; }
    };
    const errNum = errOf('StrReverse(Null)');
    assert.strictEqual(errNum !== null, true, 'StrReverse(Null) は VBA エラーを投げる');
    console.log('[PASS] Bug CZ: StrReverse(Null) はエラー (Error', errNum, ')');
}

console.log('\n✅ StrReverse: 全テスト通過');
