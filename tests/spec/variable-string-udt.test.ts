import { strict as assert } from 'node:assert';
import { evalVBASingle } from '../../test-libs/test-runner';

const source = String.raw`
Type VariableRecord
    Id As Long
    Name As String
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
`;
const ev = evalVBASingle(source);
assert.equal(ev.callProcedure('Probe', []), '42:Aあ');
console.log('[PASS] variable-string UDT Put/Get boundary (1 case)');
