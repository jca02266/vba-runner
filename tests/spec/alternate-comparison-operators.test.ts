import { evalVBASingle, assert } from '../../test-libs/test-runner';

const ev = evalVBASingle(`
    Function TestAlternateOperators() As String
        Dim result As String
        If 3 >< 5 Then result = result & "A"
        If 3 =< 3 Then result = result & "B"
        If 3 => 3 Then result = result & "C"
        TestAlternateOperators = result
    End Function

    Function TestCaseIsAlternateOperator() As String
        Select Case 4
            Case Is >< 5
                TestCaseIsAlternateOperator = "different"
            Case Else
                TestCaseIsAlternateOperator = "same"
        End Select
    End Function
`);

assert.strictEqual(ev.callProcedure('TestAlternateOperators', []), 'ABC');
assert.strictEqual(ev.callProcedure('TestCaseIsAlternateOperator', []), 'different');

console.log('✅ VBA alternate comparison operator spellings work');
