import { evalVBASingle, assert } from '../../test-libs/test-runner';

function runFunc(code: string, name: string, args: any[] = []): any {
    const ev = evalVBASingle(code, { onPrint: () => {} });
    return ev.callProcedure(name, args);
}

// 1. On...GoTo Statement
{
    const code = `
        Function TestOnGoTo(idx)
            On idx GoTo Label1, Label2, Label3
            TestOnGoTo = 0
            Exit Function
        Label1:
            TestOnGoTo = 1
            Exit Function
        Label2:
            TestOnGoTo = 2
            Exit Function
        Label3:
            TestOnGoTo = 3
        End Function
    `;
    assert.strictEqual(runFunc(code, 'TestOnGoTo', [1]), 1, 'On 1 GoTo Label1');
    assert.strictEqual(runFunc(code, 'TestOnGoTo', [2]), 2, 'On 2 GoTo Label2');
    assert.strictEqual(runFunc(code, 'TestOnGoTo', [3]), 3, 'On 3 GoTo Label3');
    assert.strictEqual(runFunc(code, 'TestOnGoTo', [0]), 0, 'On 0 GoTo should skip');
    assert.strictEqual(runFunc(code, 'TestOnGoTo', [4]), 0, 'On 4 GoTo should skip');
}

// 2. GoSub and Return
{
    const code = `
        Function TestGoSub()
            Dim x
            x = 10
            GoSub MySub
            x = x + 1
            TestGoSub = x
            Exit Function
        MySub:
            x = x * 2
            Return
        End Function
    `;
    assert.strictEqual(runFunc(code, 'TestGoSub'), 21, 'GoSub should jump and Return should come back');
}

// 3. Nested GoSub
{
    const code = `
        Function TestNestedGoSub()
            Dim x
            x = 1
            GoSub Sub1
            TestNestedGoSub = x
            Exit Function
        Sub1:
            x = x + 10
            GoSub Sub2
            Return
        Sub2:
            x = x * 2
            Return
        End Function
    `;
    assert.strictEqual(runFunc(code, 'TestNestedGoSub'), 22, 'Nested GoSub should work with LIFO stack');
}

// 4. On...GoSub
{
    const code = `
        Function TestOnGoSub(idx)
            Dim x
            x = 100
            On idx GoSub S1, S2
            TestOnGoSub = x
            Exit Function
        S1:
            x = x + 1
            Return
        S2:
            x = x + 2
            Return
        End Function
    `;
    assert.strictEqual(runFunc(code, 'TestOnGoSub', [1]), 101, 'On 1 GoSub S1');
    assert.strictEqual(runFunc(code, 'TestOnGoSub', [2]), 102, 'On 2 GoSub S2');
}

// 5. GoSub/Return inside repeated and conditional blocks must resume the
// statement after GoSub, not the statement after the containing block.
{
    const code = `
        Function TestLoopGoSub() As String
            Dim i As Long, result As String
            For i = 1 To 3
                result = result & CStr(i)
                GoSub Part
            Next i
            TestLoopGoSub = result
            Exit Function
        Part:
            result = result & "P"
            Return
        End Function
    `;
    assert.strictEqual(runFunc(code, 'TestLoopGoSub'), '1P2P3P', 'For GoSub resumes each iteration');
}
{
    const code = `
        Function TestConditionalGoSub(flag As Boolean) As String
            Dim result As String
            If flag Then
                GoSub Part
                result = result & "A"
            End If
            result = result & "B"
            TestConditionalGoSub = result
            Exit Function
        Part:
            result = result & "P"
            Return
        End Function
    `;
    assert.strictEqual(runFunc(code, 'TestConditionalGoSub', [true]), 'PAB', 'If GoSub resumes within the clause');
}
{
    const code = `
        Function TestForEachGoSub() As String
            Dim values(1 To 3) As Long, item As Variant, result As String
            values(1) = 1: values(2) = 2: values(3) = 3
            For Each item In values
                result = result & CStr(item)
                GoSub Part
            Next item
            TestForEachGoSub = result
            Exit Function
        Part:
            result = result & "P"
            Return
        End Function
    `;
    assert.strictEqual(runFunc(code, 'TestForEachGoSub'), '1P2P3P', 'For Each GoSub resumes each element');
}
{
    const code = `
        Function TestDoGoSub() As String
            Dim i As Long, result As String
            i = 1
            Do While i <= 3
                result = result & CStr(i)
                GoSub Part
                i = i + 1
            Loop
            TestDoGoSub = result
            Exit Function
        Part:
            result = result & "P"
            Return
        End Function
    `;
    assert.strictEqual(runFunc(code, 'TestDoGoSub'), '1P2P3P', 'Do While GoSub resumes the loop');
}
{
    const code = `
        Function TestSelectGoSub() As String
            Dim result As String
            Select Case 1
                Case 1
                    GoSub Part
                    result = result & "A"
            End Select
            result = result & "B"
            TestSelectGoSub = result
            Exit Function
        Part:
            result = result & "P"
            Return
        End Function
    `;
    assert.strictEqual(runFunc(code, 'TestSelectGoSub'), 'PAB', 'Select GoSub resumes within the clause');
}
{
    const code = `
        Function TestProcedureBoundary() As String
            GoSub Part
            TestProcedureBoundary = "after"
            Exit Function
        Part:
            Call HelperWithBareReturn
            Return
        End Function
        Sub HelperWithBareReturn()
            Return
        End Sub
    `;
    const ev = evalVBASingle(code, { onPrint: () => {} });
    let error: any;
    try {
        ev.callProcedure('TestProcedureBoundary', []);
    } catch (e: any) {
        error = e;
    }
    assert.strictEqual(error?.number, 3, 'Return in another procedure does not consume caller GoSub');
}

console.log('\n✅ On...GoTo/GoSub & GoSub/Return: 全テスト通過');
