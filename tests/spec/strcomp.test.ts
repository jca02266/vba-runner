import { evalVBASingle, assert, vbaNull } from '../../test-libs/test-runner';

function evalVBA(code: string): any {
    return evalVBASingle(code);
}

function runFunc(code: string, name: string, args: any[] = []): any {
    return evalVBA(code).callProcedure(name, args);
}

// StrComp のテスト
{
    const code = `
        Function TestEqual()
            TestEqual = StrComp("abc", "abc")
        End Function
        Function TestLess()
            TestLess = StrComp("abc", "def")
        End Function
        Function TestGreater()
            TestGreater = StrComp("def", "abc")
        End Function
        Function TestCaseSensitive()
            TestCaseSensitive = StrComp("abc", "ABC")
        End Function
        Function TestCaseInsensitive()
            TestCaseInsensitive = StrComp("abc", "ABC", 1) ' vbTextCompare
        End Function
    `;
    assert.strictEqual(runFunc(code, 'TestEqual'), 0, 'StrComp: "abc" = "abc" -> 0');
    assert.strictEqual(runFunc(code, 'TestLess'), -1, 'StrComp: "abc" < "def" -> -1');
    assert.strictEqual(runFunc(code, 'TestGreater'), 1, 'StrComp: "def" > "abc" -> 1');
    assert.strictEqual(runFunc(code, 'TestCaseSensitive'), 1, 'StrComp: "abc" > "ABC" in binary compare (a=97, A=65)');
    assert.strictEqual(runFunc(code, 'TestCaseInsensitive'), 0, 'StrComp: "abc" = "ABC" in text compare -> 0');
    console.log('[PASS] 基本動作');
}

// エッジケース
{
    const code = `
        Function TestNull1()
            TestNull1 = StrComp(Null, "abc")
        End Function
        Function TestNull2()
            TestNull2 = StrComp("abc", Null)
        End Function
        Function TestEmpty()
            TestEmpty = StrComp("", "")
        End Function
    `;
    assert.strictEqual(runFunc(code, 'TestNull1'), vbaNull, 'StrComp: Null input -> Null');
    assert.strictEqual(runFunc(code, 'TestNull2'), vbaNull, 'StrComp: Null input -> Null');
    assert.strictEqual(runFunc(code, 'TestEmpty'), 0, 'StrComp: "" = "" -> 0');
    console.log('[PASS] エッジケース');
}

// Compare=Null → error (spec §6.1.2.11.1.36: "If Compare is Null, an error occurs")
{
    const ev = evalVBASingle('');
    assert.throwsMatch(
        () => ev.evalExpression('StrComp("a", "b", Null)'),
        /Invalid use of Null|Null/,
        'StrComp with Compare=Null should error',
    );
    console.log('[PASS] Compare=Null はエラー');
}

console.log('\n✅ StrComp: 全テスト通過');
