import { strict as assert } from 'node:assert';
import { evalVBASingle } from '../../test-libs/test-runner';

const source = String.raw`
Type VariableRecord
    Id As Long
    Name As String
End Type

Type VariableArrayRecord
    Names(0 To 1) As String
End Type

Type MixedRecord
    Names(0 To 1) As String
    Values(0 To 1) As Long
End Type

Public Function Probe() As String
    Dim written As VariableRecord, readBack As VariableRecord
    written.Id = 42
    written.Name = "Aあ"
    Open "C:\\variable-udt.bin" For Binary As #1
    Put #1, , written
    Close #1
    Open "C:\\variable-udt.bin" For Binary As #1
    Get #1, , readBack
    Close #1
    Probe = CStr(readBack.Id) & ":" & readBack.Name
End Function

Public Function ProbeArray() As String
    Dim written As VariableArrayRecord, readBack As VariableArrayRecord
    written.Names(0) = "A"
    written.Names(1) = "あ"
    Open "C:\\variable-array-udt.bin" For Binary As #1
    Put #1, , written
    Close #1
    Open "C:\\variable-array-udt.bin" For Binary As #1
    Get #1, , readBack
    Close #1
    ProbeArray = readBack.Names(0) & ":" & readBack.Names(1)
End Function

Public Function ProbeMixed() As String
    Dim written As MixedRecord, readBack As MixedRecord
    written.Names(0) = "A"
    written.Names(1) = "あ"
    written.Values(0) = 10
    written.Values(1) = 20
    Open "C:\\mixed-udt.bin" For Binary As #1
    Put #1, , written
    Close #1
    Open "C:\\mixed-udt.bin" For Binary As #1
    Get #1, , readBack
    Close #1
    ProbeMixed = readBack.Names(0) & ":" & readBack.Names(1) & ":" & _
        CStr(readBack.Values(0)) & ":" & CStr(readBack.Values(1))
End Function
`;
const ev = evalVBASingle(source);
assert.equal(ev.callProcedure('Probe', []), '42:Aあ');
assert.equal(ev.callProcedure('ProbeArray', []), 'A:あ');
assert.equal(ev.callProcedure('ProbeMixed', []), 'A:あ:10:20');
console.log('[PASS] variable-string UDT Put/Get boundary (3 cases)');
