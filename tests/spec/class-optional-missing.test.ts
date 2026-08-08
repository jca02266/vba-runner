import { evalVBASingle, assert } from '../../test-libs/test-runner';

const code = String.raw`
Class Payload
    Public Property Get Value() As Long
        Value = 4
    End Property
End Class

Class Box
    Private letSeen As String
    Private setSeen As String

    Public Function Sum(a As Long, Optional middle As Long = 7, Optional tail As Long = 9) As Long
        Sum = a + middle + tail
    End Function

    Public Property Get Total(a As Long, Optional middle As Long = 7, Optional tail As Long = 9) As Long
        Total = a + middle + tail
    End Property

    Public Property Let Label(a As Long, Optional middle As Long = 7, Optional tail As Long = 9, value As String)
        letSeen = CStr(a) & ":" & CStr(middle) & ":" & CStr(tail) & ":" & value
    End Property

    Public Property Set Item(a As Long, Optional middle As Long = 7, Optional tail As Long = 9, value As Object)
        setSeen = CStr(a) & ":" & CStr(middle) & ":" & CStr(tail) & ":" & CStr(value.Value)
    End Property

    Public Property Get Result() As String
        Result = letSeen & "|" & setSeen
    End Property
End Class

Function Probe() As String
    Dim box As New Box, payload As New Payload
    box.Label 1,,3,"ok"
    Set box.Item(1,,3) = payload
    Probe = CStr(box.Sum(1,,3)) & "|" & CStr(box.Total(1,,3)) & "|" & box.Result
End Function
`;

const evaluator = evalVBASingle(code);
assert.strictEqual(evaluator.callProcedure('Probe', []), '11|11|1:7:3:ok|1:7:3:4',
    'Class Function and Property procedures apply Optional defaults to omitted slots');
console.log('[PASS] Class Optional MissingArgument defaults');
