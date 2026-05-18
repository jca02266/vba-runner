import { Lexer } from '../../src/engine/lexer';
import { Parser } from '../../src/engine/parser';
import { Evaluator } from '../../src/engine/evaluator';
import { assert } from '../../test-libs/test-runner';

function evalVBA(code: string): any {
    const tokens = new Lexer(code).tokenize();
    const ast = new Parser(tokens).parse();
    const ev = new Evaluator(console.log);
    ev.evaluate(ast);
    return ev;
}

function runFunc(code: string, name: string, args: any[] = []): any {
    return evalVBA(code).callProcedure(name, args);
}

// Test 1: On Error Resume Next inside Select Case catches error at statement level
{
    const result = runFunc(`
Function F() As String
    Dim lb As Long
    Dim lb2 As Long
    lb = -1
    lb2 = -1
    Dim arr(1 To 3) As Long
    arr(1) = 10
    arr(2) = 20
    arr(3) = 30
    Select Case VarType(arr)
    Case 8192 To 8209
        On Error Resume Next
        lb = LBound(arr, 1)
        lb2 = LBound(arr, 2)
        On Error GoTo 0
        F = "lb=" & lb & " lb2=" & lb2
    Case Else
        F = "not array"
    End Select
End Function
`, 'F');
    assert.strictEqual(result, 'lb=1 lb2=-1', 'On Error Resume Next inside Select Case: bad LBound leaves var unchanged');
    console.log('[PASS] On Error Resume Next inside Select Case');
}

// Test 2: On Error Resume Next inside If block
{
    const result = runFunc(`
Function F() As String
    Dim x As Long
    x = 0
    Dim y As Long
    y = -1
    If x = 0 Then
        On Error Resume Next
        y = 10 / x
        On Error GoTo 0
        F = "y=" & y
    End If
End Function
`, 'F');
    assert.strictEqual(result, 'y=-1', 'On Error Resume Next inside If body: error skips statement');
    console.log('[PASS] On Error Resume Next inside If block');
}

// Test 3: On Error Resume Next inside For loop
{
    const result = runFunc(`
Function F() As Long
    Dim acc As Long
    acc = 0
    Dim i As Long
    On Error Resume Next
    For i = -1 To 1
        acc = acc + (10 \\ i)
    Next i
    On Error GoTo 0
    F = acc
End Function
`, 'F');
    // i=-1: 10 \\ -1 = -10; i=0: div by zero caught, skipped; i=1: 10 \\ 1 = 10 → total 0
    assert.strictEqual(result, 0, 'On Error Resume Next in For body: div by zero on i=0 skipped, total=-10+0+10=0');
    console.log('[PASS] On Error Resume Next inside For loop');
}

// Test 4: On Error Resume Next inside While loop
{
    const result = runFunc(`
Function F() As Long
    Dim total As Long
    total = 0
    Dim x As Long
    x = 0
    On Error Resume Next
    While x <= 2
        If x > 0 Then
            total = total + (10 \\ x)
        End If
        x = x + 1
    Wend
    On Error GoTo 0
    F = total
End Function
`, 'F');
    // x=0: skip (x>0 is false); x=1: 10\\1=10; x=2: 10\\2=5 → total=15
    assert.strictEqual(result, 15, 'On Error Resume Next inside While: 10/1 + 10/2 = 15');
    console.log('[PASS] On Error Resume Next inside While loop');
}

// Test 5: Resume in On Error GoTo handler inside nested If block propagates correctly
{
    const result = runFunc(`
Function F() As Long
    Dim x As Long
    Dim retried As Boolean
    x = 0
    retried = False
    On Error GoTo Handler
    F = 10 / x
    Exit Function
Handler:
    If Not retried Then
        retried = True
        x = 2
        Resume
    End If
    F = 999
End Function
`, 'F');
    assert.strictEqual(result, 5, 'Resume in GoTo handler inside If block: retries with x=2, returns 5');
    console.log('[PASS] Resume inside If block in error handler');
}

// Test 6: On Error Resume Next across nested Select Case + loop pattern
{
    const result = runFunc(`
Function F() As String
    Dim arr(1 To 3) As Variant
    arr(1) = 1
    arr(2) = "two"
    arr(3) = True

    Dim lb As Long
    Dim ub As Long
    Dim lb2 As Long
    Dim ub2 As Long
    lb = -1
    ub = -1
    lb2 = -1
    ub2 = -1

    Dim result As String
    result = ""

    Select Case VarType(arr)
    Case 8192 To 8209
        On Error Resume Next
        lb = LBound(arr, 1)
        ub = UBound(arr, 1)
        lb2 = LBound(arr, 2)
        ub2 = UBound(arr, 2)
        On Error GoTo 0

        Dim j As Long
        If lb >= 0 And ub >= 0 Then
            If lb2 < 0 Or ub2 < 0 Then
                For j = lb To ub
                    result = result & CStr(arr(j)) & ","
                Next j
            End If
        End If
    End Select

    F = result
End Function
`, 'F');
    assert.strictEqual(result, '1,two,True,', 'On Error Resume Next in Select+For: 1D array iteration with 2D boundary error');
    console.log('[PASS] On Error Resume Next in Select Case + nested For pattern');
}

console.log('\n✅ on-error-in-blocks: 全テスト通過');
