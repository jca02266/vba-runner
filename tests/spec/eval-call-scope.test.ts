/**
 * evalExpression 経由の関数呼び出しにおけるスコープ分離のテスト
 *
 * バグ修正: evaluator.ts の CallExpression インライン経路が
 * currentProcedureName を設定していなかったため、callee の Dim 文が
 * enclosing を辿る set() を使い、呼び出し元の同名変数を破壊していた。
 * 内外で同名ループ変数を持つネスト呼び出しで無限ループが発生していた。
 */
import { evalVBASingle, assert } from '../../test-libs/test-runner';

function run(code: string, name: string, args: any[] = []): any {
    return evalVBASingle(code).callProcedure(name, args);
}

// -----------------------------------------------------------------------
// 1. 内外同名変数 i を持つネスト For ループ — eval 経由
//    callee の Dim i が caller のループ変数 i を破壊しないこと
// -----------------------------------------------------------------------
{
    const code = `
Function Outer() As String
    Dim i As Long
    Dim result As String
    For i = 1 To 4 Step 2
        result = result & Inner(i) & ","
    Next i
    Outer = result
End Function

Function Inner(n As Long) As Long
    Dim i As Long
    Dim total As Long
    For i = 1 To n
        total = total + i
    Next i
    Inner = total
End Function
`;
    // i=1: Inner(1)=1, i=3: Inner(3)=6 → "1,6,"
    const ev = evalVBASingle(code);
    assert.strictEqual(ev.evalExpression('Outer()'), '1,6,',
        'eval経由: 内外同名変数 i を持つネストFor — 呼び出し元ループ変数が保持される');
    assert.strictEqual(ev.callProcedure('Outer', []), '1,6,',
        'run経由: 同上（リグレッション）');
}
console.log('[PASS] 内外同名変数 i のネスト For ループ');

// -----------------------------------------------------------------------
// 2. HexDecode パターン — Step 2 の For ループ内でヘルパーを呼ぶ
//    元のバグ再現ケース（eval経由でハングしていた）
// -----------------------------------------------------------------------
{
    const code = `
Function HexDecode(hexStr As String) As String
    Dim result As String
    Dim i As Long
    For i = 1 To Len(hexStr) Step 2
        Dim pair As String
        pair = Mid(hexStr, i, 2)
        If Not IsHexPair(pair) Then
            Err.Raise 13, "HexDecode", "invalid hex"
        End If
        result = result & Chr(HexCharToNum(Left(pair, 1)) * 16 + HexCharToNum(Right(pair, 1)))
    Next i
    HexDecode = result
End Function

Private Function IsHexPair(s As String) As Boolean
    Dim i As Long
    IsHexPair = True
    For i = 1 To Len(s)
        Dim c As String
        c = UCase(Mid(s, i, 1))
        If Not ((c >= "0" And c <= "9") Or (c >= "A" And c <= "F")) Then
            IsHexPair = False
            Exit Function
        End If
    Next i
End Function

Private Function HexCharToNum(c As String) As Integer
    Dim code As Integer
    code = Asc(UCase(c))
    If code >= 48 And code <= 57 Then
        HexCharToNum = code - 48
    ElseIf code >= 65 And code <= 70 Then
        HexCharToNum = code - 55
    End If
End Function
`;
    const ev = evalVBASingle(code);
    assert.strictEqual(ev.evalExpression('HexDecode("4142")'), 'AB',
        'HexDecode("4142") via eval => "AB"');
    assert.strictEqual(ev.evalExpression('HexDecode("48656C6C6F")'), 'Hello',
        'HexDecode("48656C6C6F") via eval => "Hello"');
    assert.strictEqual(ev.evalExpression('HexDecode("")'), '',
        'HexDecode("") via eval => ""');
    assert.strictEqual(ev.callProcedure('HexDecode', ['4142']), 'AB',
        'HexDecode("4142") via run => "AB"（リグレッション）');
}
console.log('[PASS] HexDecode パターン（Step 2 ループ + 内側ヘルパー呼び出し）');

console.log('\n✅ eval-call-scope: 全テスト通過');
