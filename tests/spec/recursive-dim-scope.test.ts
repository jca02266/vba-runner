import { evalVBASingle, assert } from '../../test-libs/test-runner';

function evalVBA(code: string): any {
    return evalVBASingle(code);
}

function runFunc(code: string, name: string, args: any[] = []): any {
    return evalVBA(code).callProcedure(name, args);
}

// Test 1: 再帰関数で outer の Dim 変数が内側の Dim に上書きされない
// バグ再現: Dim result As String を持つ関数が自分自身を再帰呼び出しすると、
// 内側の Dim が外側の result を "" に reset していた
{
    const result = runFunc(`
Function Concat(n As Long) As String
    Dim result As String
    If n <= 0 Then
        Concat = ""
        Exit Function
    End If
    result = "[" & n & "]"
    result = result & Concat(n - 1)
    Concat = result
End Function
`, 'Concat', [3]);
    assert.strictEqual(result, '[3][2][1]', 'Recursive function: outer Dim not overwritten by inner Dim');
    console.log('[PASS] Recursive function: outer Dim not overwritten');
}

// Test 2: 相互再帰（2関数が互いを呼ぶ）でもスコープ独立
{
    const result = runFunc(`
Function Even(n As Long) As String
    Dim tag As String
    tag = "E" & n
    If n = 0 Then
        Even = tag
        Exit Function
    End If
    Even = tag & Odd(n - 1)
End Function

Function Odd(n As Long) As String
    Dim tag As String
    tag = "O" & n
    If n = 0 Then
        Odd = tag
        Exit Function
    End If
    Odd = tag & Even(n - 1)
End Function
`, 'Even', [3]);
    assert.strictEqual(result, 'E3O2E1O0', 'Mutual recursion: each call has its own Dim scope');
    console.log('[PASS] Mutual recursion: each call has its own Dim scope');
}

// Test 3: 再帰でカウンタ系 Dim が正しく動く（フィボナッチ）
{
    const result = runFunc(`
Function Fib(n As Long) As Long
    Dim a As Long
    Dim b As Long
    If n <= 1 Then
        Fib = n
        Exit Function
    End If
    a = Fib(n - 1)
    b = Fib(n - 2)
    Fib = a + b
End Function
`, 'Fib', [7]);
    assert.strictEqual(result, 13, 'Fibonacci(7) = 13: recursive Dim a/b stay local');
    console.log('[PASS] Fibonacci(7) = 13: recursive Dim a/b stay local');
}

// Test 4: ByRef Sub と再帰関数が混在しても独立
// json_BufferAppend パターン: ByRef 文字列バッファを使う再帰的な変換関数
{
    const result = runFunc(`
Sub BufAdd(ByRef buf As String, ByVal s As String)
    buf = buf & s
End Sub

Function Serialize(n As Long) As String
    Dim buf As String
    Dim inner As String
    If n = 0 Then
        Serialize = "0"
        Exit Function
    End If
    BufAdd buf, "("
    inner = Serialize(n - 1)
    BufAdd buf, inner
    BufAdd buf, ")"
    Serialize = buf
End Function
`, 'Serialize', [3]);
    assert.strictEqual(result, '(((0)))', 'ByRef Sub + recursive: buf stays local per call');
    console.log('[PASS] ByRef Sub + recursive: buf stays local per call');
}

// Test 5: 同名の Dim を持つ非再帰 2 関数は互いに干渉しない
{
    const result = runFunc(`
Function A() As String
    Dim x As String
    x = "from_A"
    A = x & B()
End Function

Function B() As String
    Dim x As String
    x = "from_B"
    B = x
End Function
`, 'A');
    assert.strictEqual(result, 'from_Afrom_B', 'Same Dim name in different functions: no interference');
    console.log('[PASS] Same Dim name in different functions: no interference');
}

console.log('\n✅ recursive-dim-scope: 全テスト通過');
