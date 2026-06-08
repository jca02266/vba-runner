import { evalVBASingle, assert } from '../../test-libs/test-runner';

function evalVBA(code: string): any {
    return evalVBASingle(code);
}

function runFunc(code: string, name: string, args: any[] = []): any {
    return evalVBA(code).callProcedure(name, args);
}

// Test 1: F = F & "X" — 戻り値変数の右辺参照が再帰呼び出しにならない（バグ再現）
{
    const result = runFunc(`
Function F() As String
    F = F & "X"
End Function
`, 'F');
    assert.strictEqual(result, 'X', 'F = F & "X": bare F on RHS reads return var, not recursive call');
    console.log('[PASS] F = F & "X" (no recursion)');
}

// Test 2: As String 関数の戻り値変数初期値は ""
{
    const result = runFunc(`
Function F() As String
    Dim s As String
    s = "AB"
    Dim i As Long
    For i = 1 To Len(s)
        F = F & Mid$(s, i, 1)
    Next i
End Function
`, 'F');
    assert.strictEqual(result, 'AB', 'String function return var accumulates correctly');
    console.log('[PASS] F = F & Mid$ loop: "AB"');
}

// Test 3: As Long 関数の戻り値変数は 0 で始まり累積できる
{
    const result = runFunc(`
Function F() As Long
    Dim i As Long
    For i = 1 To 5
        F = F + i
    Next i
End Function
`, 'F');
    assert.strictEqual(result, 15, 'Long function: F = F + i accumulates 1+2+3+4+5 = 15');
    console.log('[PASS] F = F + i loop: 15');
}

// Test 4: 同名の別関数を呼ぶ（G の中から F を呼ぶ → auto-call すべき）
{
    const result = runFunc(`
Function F() As String
    F = "from_F"
End Function

Function G() As String
    G = F & "_from_G"
End Function
`, 'G');
    assert.strictEqual(result, 'from_F_from_G', 'Inside G, bare F is auto-called as 0-arg function');
    console.log('[PASS] G calls F auto-call');
}

// Test 5: VBA.InStr + VBA.Mid$ によるキャラクタスキャンパターン
{
    const result = runFunc(`
Function F() As String
    Dim s As String
    s = "hello123world"
    Dim acc As String
    Dim i As Long
    For i = 1 To Len(s)
        If VBA.InStr("0123456789", VBA.Mid$(s, i, 1)) > 0 Then
            acc = acc & VBA.Mid$(s, i, 1)
        End If
    Next i
    F = acc
End Function
`, 'F');
    assert.strictEqual(result, '123', 'VBA.InStr + VBA.Mid$ char scan extracts digits');
    console.log('[PASS] VBA.InStr + VBA.Mid$ digit scan: "123"');
}

// Test 6: For Each In Collection を返す Function から
{
    const result = runFunc(`
Function MakeCol() As Collection
    Dim c As New Collection
    c.Add "x"
    c.Add "y"
    c.Add "z"
    Set MakeCol = c
End Function

Function F() As String
    Dim acc As String
    Dim v As Variant
    For Each v In MakeCol()
        acc = acc & v
    Next v
    F = acc
End Function
`, 'F');
    assert.strictEqual(result, 'xyz', 'For Each in Collection returned from function');
    console.log('[PASS] For Each in Collection from Function: "xyz"');
}

console.log('\n✅ function-return-var: 全テスト通過');
