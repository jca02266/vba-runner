import { strict as assert } from 'node:assert';
import { evalVBASingle } from '../../test-libs/test-runner';

const source = String.raw`
Type NestedChild
    Id As Long
    Code As String * 2
End Type

Type NestedParent
    Header As NestedChild
    Values(0 To 1) As Integer
End Type

Public Function Probe() As String
    Dim written As NestedParent, readBack As NestedParent
    written.Header.Id = 42
    written.Header.Code = "XY"
    written.Values(0) = 100
    written.Values(1) = 200
    Open "C:\\nested-udt.bin" For Binary As #1
    Put #1, , written
    Close #1
    Open "C:\\nested-udt.bin" For Binary As #1
    Get #1, , readBack
    Close #1
    Probe = CStr(readBack.Header.Id) & ":" & readBack.Header.Code & ":" & _
        CStr(readBack.Values(0)) & ":" & CStr(readBack.Values(1))
End Function
`;
const ev = evalVBASingle(source);
assert.equal(ev.callProcedure('Probe', []), '42:XY:100:200');
console.log('[PASS] nested UDT array Put/Get boundary (1 case)');
