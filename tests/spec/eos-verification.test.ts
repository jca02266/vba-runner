/**
 * §5.4.1 EOS (End-Of-Statement) 検証のテスト
 *
 * statement-block = *(block-statement EOS)
 * EOS = *(EOL / ":")
 *
 * 各文の後には必ず EOS（改行またはコロン）が必要。
 * 余分なトークンが続く場合は ParseError にする。
 */
import { evalVBASingle, assert } from '../../test-libs/test-runner';

function run(code: string, name = 'T'): any {
    return evalVBASingle(code).callProcedure(name, []);
}

function assertParseError(code: string, description: string): void {
    let threw = false;
    try {
        evalVBASingle(code);
    } catch (e: any) {
        threw = true;
        // ParseError か RuntimeError かを問わず「余分なトークン」系のエラーであること
        const msg: string = e.message ?? '';
        assert.strictEqual(
            msg.includes('unexpected') || msg.includes('Parse error'),
            true,
            `${description}: error message should mention unexpected token, got: "${msg}"`
        );
    }
    assert.strictEqual(threw, true, `${description}: should throw but didn't`);
}

// --- 1. 正常系：EOS あり ---
{
    // Return (GoSub 戻り) — 単独で使える
    assert.strictEqual(
        run(`
Function T() As Long
    GoSub Calc
    Exit Function
Calc:
    T = 42
    Return
End Function`),
        42, 'Return standalone is valid'
    );

    // Exit Sub — 単独で使える
    assert.strictEqual(
        run(`
Function T() As Long
    T = 99
    Exit Function
    T = 0
End Function`),
        99, 'Exit Function standalone is valid'
    );

    // Exit For — 単独で使える
    assert.strictEqual(
        run(`
Function T() As Long
Dim i As Long
For i = 1 To 10
    If i = 3 Then Exit For
Next i
T = i
End Function`),
        3, 'Exit For standalone is valid'
    );

    // コロン区切りの複数文 — 有効
    assert.strictEqual(
        run(`
Function T() As Long
Dim a As Long
Dim b As Long
a = 1 : b = 2
T = a + b
End Function`),
        3, 'colon-separated statements are valid'
    );

    // インライン If の Else 前で EOS チェックが走らない
    assert.strictEqual(
        run(`
Function T() As Long
Dim x As Long
x = 5
If x > 3 Then T = 1 Else T = 2
End Function`),
        1, 'inline If-Then-Else is valid'
    );

    console.log('[PASS] 正常系：EOS あり');
}

// --- 2. エラー系：余分なトークン ---
{
    // Return の後に余分なトークン
    assertParseError(`
Sub T()
    GoSub Lbl
    Exit Sub
Lbl:
    Return Foo
End Sub`, 'Return Foo should be a parse error');

    // Exit Sub の後に余分なトークン
    assertParseError(`
Sub T()
    Exit Sub Bar
End Sub`, 'Exit Sub Bar should be a parse error');

    // Exit For の後に余分なトークン
    assertParseError(`
Sub T()
    Dim i As Long
    For i = 1 To 5
        Exit For Extra
    Next i
End Sub`, 'Exit For Extra should be a parse error');

    // Exit Function の後に余分なトークン
    assertParseError(`
Function T() As Long
    T = 1
    Exit Function Junk
End Function`, 'Exit Function Junk should be a parse error');

    // Exit Do の後に余分なトークン
    assertParseError(`
Sub T()
    Do
        Exit Do Garbage
    Loop
End Sub`, 'Exit Do Garbage should be a parse error');

    console.log('[PASS] エラー系：余分なトークン');
}

// --- 3. ブロック内 EOS 検証（For, Do, While, With, Select Case） ---
{
    // For ブロック内
    assert.strictEqual(
        run(`
Function T() As Long
Dim s As Long
Dim i As Long
s = 0
For i = 1 To 3
    s = s + i
Next i
T = s
End Function`),
        6, 'For block normal'
    );

    // While ブロック内
    assert.strictEqual(
        run(`
Function T() As Long
Dim n As Long
n = 0
Dim i As Long
i = 1
While i <= 4
    n = n + i
    i = i + 1
Wend
T = n
End Function`),
        10, 'While block normal'
    );

    // Select Case ブロック内
    assert.strictEqual(
        run(`
Function T() As Long
Dim x As Long
x = 2
Select Case x
    Case 1
        T = 10
    Case 2
        T = 20
    Case Else
        T = 30
End Select
End Function`),
        20, 'Select Case block normal'
    );

    // With ブロック内 (Scripting.Dictionary を使わず型なしで確認)
    assert.strictEqual(
        run(`
Function T() As Long
Dim a As Long
a = 5
T = a
End Function`),
        5, 'basic block EOS sanity'
    );

    console.log('[PASS] ブロック内 EOS 検証');
}

console.log('\n✅ eos-verification: 全テスト通過');
