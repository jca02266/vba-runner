import { evalVBASingle, assert, vbaNull } from '../../test-libs/test-runner';

function evalVBA(code: string): any {
    return evalVBASingle(code);
}

function runFunc(code: string, name: string, args: any[] = []): any {
    return evalVBA(code).callProcedure(name, args);
}

// 基本動作
{
    const code = `
        Function Test1()
            Test1 = InStrRev("alphabet", "a")
        End Function
        Function Test2()
            Test2 = InStrRev("alphabet", "a", 5)
        End Function
        Function Test3()
            Test3 = InStrRev("alphabet", "z")
        End Function
        Function Test4()
            Test4 = InStrRev("alphabet", "A")
        End Function
        Function Test5()
            Test5 = InStrRev("alphabet", "A", -1, 1) ' vbTextCompare
        End Function
    `;
    assert.strictEqual(runFunc(code, 'Test1'), 5, 'InStrRev: last "a" in "alphabet" should be at 5');
    // Wait, "alphabet" has "a" at 1. 
    // "banana" -> InStrRev("banana", "a") -> 6
    
    const code2 = `
        Function Test6()
            Test6 = InStrRev("banana", "a")
        End Function
        Function Test7()
            Test7 = InStrRev("banana", "a", 5)
        End Function
    `;
    assert.strictEqual(runFunc(code2, 'Test6'), 6, 'InStrRev: "banana", "a" -> 6');
    assert.strictEqual(runFunc(code2, 'Test7'), 4, 'InStrRev: "banana", "a", 5 -> 4');
    
    assert.strictEqual(runFunc(code, 'Test3'), 0, 'InStrRev: not found -> 0');
    assert.strictEqual(runFunc(code, 'Test4'), 0, 'InStrRev: case sensitive by default');
    assert.strictEqual(runFunc(code, 'Test5'), 5, 'InStrRev: case insensitive with vbTextCompare');
    console.log('[PASS] 基本動作');
}

// エッジケース
{
    const code = `
        Function TestNull1()
            TestNull1 = InStrRev(Null, "a")
        End Function
        Function TestNull2()
            TestNull2 = InStrRev("a", Null)
        End Function
        Function TestEmpty1()
            TestEmpty1 = InStrRev("", "a")
        End Function
        Function TestEmpty2()
            TestEmpty2 = InStrRev("abc", "")
        End Function
        Function TestStartOver()
            TestStartOver = InStrRev("abc", "a", 10)
        End Function
    `;
    assert.strictEqual(runFunc(code, 'TestNull1'), vbaNull, 'InStrRev: StringCheck is Null -> Null');
    assert.strictEqual(runFunc(code, 'TestNull2'), vbaNull, 'InStrRev: StringMatch is Null -> Null');
    assert.strictEqual(runFunc(code, 'TestEmpty1'), 0, 'InStrRev: StringCheck is empty -> 0');
    // In VBA: InStrRev("abc", "") with default start (-1) returns 3 (length of StringCheck)
    assert.strictEqual(runFunc(code, 'TestEmpty2'), 3, 'InStrRev: StringMatch is empty -> Len(StringCheck)');
    assert.strictEqual(runFunc(code, 'TestStartOver'), 0, 'InStrRev: Start > Len(StringCheck) -> 0');
    console.log('[PASS] エッジケース');
}

// --- Bug CJ: Start/Compare の Null・不正値、空文字列の優先順位 (§6.1.2.11.1.15) ---
{
    const ev = (expr: string) => evalVBA('').evalExpression(expr);
    const errOf = (expr: string) => {
        try { ev(expr); return 0; } catch (e: any) { return e?.number ?? -1; }
    };

    // 仕様: "If Start contains the data value Null, an error occurs" → error 94
    assert.strictEqual(errOf('InStrRev("abc", "b", Null)'), 94, 'InStrRev: Start=Null -> error 94');
    // Compare=Null もエラー（InStr Bug BH と同様）
    assert.strictEqual(errOf('InStrRev("abc", "b", -1, Null)'), 94, 'InStrRev: Compare=Null -> error 94');
    // Start が 0 または -1 以外の負数 → error 5
    assert.strictEqual(errOf('InStrRev("abc", "a", 0)'), 5, 'InStrRev: Start=0 -> error 5');
    assert.strictEqual(errOf('InStrRev("abc", "a", -2)'), 5, 'InStrRev: Start=-2 -> error 5');
    // 仕様テーブル: StringCheck zero-length → 0（StringMatch zero-length → Start より優先）
    assert.strictEqual(ev('InStrRev("", "", 5)'), 0, 'InStrRev("", "", 5) -> 0');
    // 仕様テーブル: Start > Len(StringCheck) → 0（StringMatch zero-length でも同様）
    assert.strictEqual(ev('InStrRev("abc", "", 5)'), 0, 'InStrRev("abc", "", 5) -> 0');
    // Start <= Len のときは StringMatch zero-length → Start
    assert.strictEqual(ev('InStrRev("abc", "", 2)'), 2, 'InStrRev("abc", "", 2) -> 2');
    console.log('[PASS] Bug CJ: InStrRev Null/不正 Start・空文字列優先順位');
}

console.log('\n✅ InStrRev: 全テスト通過');
