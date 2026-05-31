import { evalVBASingle } from '../../test-libs/test-runner';
import { assert } from '../../test-libs/test-runner';

// Test 1: Dim ws As <未知の外部型> の初期値は Nothing（vbEmpty ではない）
{
    const result = evalVBASingle(`
Function F() As String
    Dim ws As Worksheet
    If ws Is Nothing Then
        F = "Nothing"
    ElseIf IsEmpty(ws) Then
        F = "Empty"
    Else
        F = "other"
    End If
End Function
`).callProcedure('F', []);
    assert.strictEqual(result, 'Nothing', 'Dim ws As Worksheet initializes to Nothing, not Empty');
    console.log('[PASS] Test 1: Dim ws As Worksheet initializes to Nothing');
}

// Test 2: On Error Resume Next で Set が失敗しても Nothing のまま
{
    const result = evalVBASingle(`
Function GetNum() As Long
    GetNum = 42
End Function

Function F() As String
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = GetNum()
    On Error GoTo 0
    If ws Is Nothing Then
        F = "Nothing"
    Else
        F = "not Nothing"
    End If
End Function
`).callProcedure('F', []);
    assert.strictEqual(result, 'Nothing', 'Set fails under On Error Resume Next, obj remains Nothing');
    console.log('[PASS] Test 2: Set fails under On Error Resume Next, obj remains Nothing');
}

// Test 3: Option Explicit でも Nothing オブジェクトへのアクセスエラー(91)は On Error Resume Next で抑制される
{
    const result = evalVBASingle(`
Option Explicit

Function F() As String
    Dim ws As Worksheet
    Dim val As Variant
    On Error Resume Next
    val = ws.Value
    On Error GoTo 0
    If IsEmpty(val) Then
        F = "empty"
    Else
        F = CStr(val)
    End If
End Function
`).callProcedure('F', []);
    assert.strictEqual(result, 'empty', 'Error 91 on Nothing object access suppressed by On Error Resume Next');
    console.log('[PASS] Test 3: Error 91 suppressed by On Error Resume Next under Option Explicit');
}

console.log('\n✅ on-error-nothing: 全テスト通過');
