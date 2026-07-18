import { evalVBASingle, assert } from '../../test-libs/test-runner';

function runFunc(code: string, name: string, args: any[] = []): any {
    const ev = evalVBASingle(code, { onPrint: () => {} });
    return ev.callProcedure(name, args);
}

// 1. LSet Statement
{
    const code = `
        Function TestLSet()
            Dim s As String
            s = "ABCDE"
            LSet s = "123"
            TestLSet = s
        End Function
    `;
    assert.strictEqual(runFunc(code, 'TestLSet'), "123  ", 'LSet shorter');
}

{
    const code = `
        Function TestLSetLonger()
            Dim s As String
            s = "ABC"
            LSet s = "12345"
            TestLSetLonger = s
        End Function
    `;
    assert.strictEqual(runFunc(code, 'TestLSetLonger'), "123", 'LSet longer');
}

// 2. RSet Statement
{
    const code = `
        Function TestRSet()
            Dim s As String
            s = "ABCDE"
            RSet s = "123"
            TestRSet = s
        End Function
    `;
    assert.strictEqual(runFunc(code, 'TestRSet'), "  123", 'RSet shorter');
}

// Bug 32-E（評価 #32）: LSet による UDT 間コピーが Error 424 だった。
// 同一レイアウト（フィールド数一致）の場合は位置ベースでフィールドをコピーする。
{
    const code = `
        Type TA
            S As String * 8
            N As Long
        End Type
        Type TB
            T As String * 8
            M As Long
        End Type
        Function TestUdtLSet() As String
            Dim a As TA, b As TB
            a.S = "HELLO"
            a.N = 7
            LSet b = a
            TestUdtLSet = b.T & "|" & b.M
        End Function
    `;
    assert.strictEqual(runFunc(code, 'TestUdtLSet'), 'HELLO   |7',
        'Bug 32-E: LSet で同一レイアウト UDT 間の位置コピーができる');
    console.log('[PASS] Bug 32-E: LSet UDT 間コピー');
}

// Bug 32-E 続: レイアウト（フィールド数）が異なる UDT 間はエラー 5 で明示的に失敗する
{
    const code = `
        Type TC
            X As Long
        End Type
        Type TD
            Y As Long
            Z As Long
        End Type
        Sub TestMismatch()
            Dim c As TC, d As TD
            LSet d = c
        End Sub
    `;
    assert.throwsMatch(() => runFunc(code, 'TestMismatch'), /error '5'/,
        'Bug 32-E: レイアウト不一致の UDT 間 LSet は Error 5');
    console.log('[PASS] Bug 32-E: レイアウト不一致 UDT の LSet は明示エラー');
}

console.log('\n✅ LSet/RSet Statement: 全テスト通過');
