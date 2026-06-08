import { evalVBASingle, assert } from '../../test-libs/test-runner';

function evalVBA(code: string): any {
    return evalVBASingle(code);
}

function runFunc(code: string, name: string, args: any[] = []): any {
    return evalVBA(code).callProcedure(name, args);
}

// Test 1: VBA.vbString 定数アクセス（バグ再現: Case VBA.vbString が失敗していた）
{
    const result = runFunc(`
Function F() As Long
    F = VBA.vbString
End Function
`, 'F');
    assert.strictEqual(result, 8, 'VBA.vbString = 8');
    console.log('[PASS] VBA.vbString = 8');
}

// Test 2: VBA.vbNull 定数アクセス
{
    const result = runFunc(`
Function F() As Long
    F = VBA.vbNull
End Function
`, 'F');
    assert.strictEqual(result, 1, 'VBA.vbNull = 1');
    console.log('[PASS] VBA.vbNull = 1');
}

// Test 3: VBA.vbBoolean 定数アクセス
{
    const result = runFunc(`
Function F() As Long
    F = VBA.vbBoolean
End Function
`, 'F');
    assert.strictEqual(result, 11, 'VBA.vbBoolean = 11');
    console.log('[PASS] VBA.vbBoolean = 11');
}

// Test 4: VBA.vbObject 定数アクセス
{
    const result = runFunc(`
Function F() As Long
    F = VBA.vbObject
End Function
`, 'F');
    assert.strictEqual(result, 9, 'VBA.vbObject = 9');
    console.log('[PASS] VBA.vbObject = 9');
}

// Test 5: Select Case で VBA. 定数を使う（バグ再現: Case VBA.vbString が実行時エラー 424 になっていた）
{
    const result = runFunc(`
Function F(x As Variant) As String
    Select Case VBA.VarType(x)
    Case VBA.vbString
        F = "string"
    Case VBA.vbBoolean
        F = "boolean"
    Case Else
        F = "other"
    End Select
End Function
`, 'F', ["hello"]);
    assert.strictEqual(result, 'string', 'Select Case VBA.vbString matches string');
    console.log('[PASS] Select Case VBA.vbString matches string');
}

// Test 6: Select Case で VBA. 定数 boolean に一致（VBA 内部の True を使う）
{
    const result = runFunc(`
Function F() As String
    Dim x As Variant
    x = True
    Select Case VBA.VarType(x)
    Case VBA.vbString
        F = "string"
    Case VBA.vbBoolean
        F = "boolean"
    Case Else
        F = "other"
    End Select
End Function
`, 'F');
    assert.strictEqual(result, 'boolean', 'Select Case VBA.vbBoolean matches VBA True');
    console.log('[PASS] Select Case VBA.vbBoolean matches VBA True');
}

// Test 7: VBA.VarType 関数呼び出し（モジュール修飾関数）
{
    const result = runFunc(`
Function F() As Long
    F = VBA.VarType("test")
End Function
`, 'F');
    assert.strictEqual(result, 8, 'VBA.VarType("test") = 8 (vbString)');
    console.log('[PASS] VBA.VarType("test") = 8');
}

// Test 8: VBA.TypeName 関数呼び出し
{
    const result = runFunc(`
Function F() As String
    F = VBA.TypeName("hello")
End Function
`, 'F');
    assert.strictEqual(result, 'String', 'VBA.TypeName("hello") = "String"');
    console.log('[PASS] VBA.TypeName("hello") = "String"');
}

console.log('\n✅ vba-module-qualifier: 全テスト通過');
