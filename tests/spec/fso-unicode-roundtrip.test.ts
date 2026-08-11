import { strict as assert } from 'node:assert';
import { evalVBASingle } from '../../test-libs/test-runner';

const source = String.raw`
Public Function Probe() As String
    Dim fso As Object, stream As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set stream = fso.CreateTextFile("C:\\unicode-roundtrip.txt", True, True)
    stream.Write "Aあ"
    stream.Close
    Set stream = fso.OpenTextFile("C:\\unicode-roundtrip.txt", 1, False, True)
    Probe = stream.ReadAll
    stream.Close
End Function
`;
const ev = evalVBASingle(source);
assert.equal(ev.callProcedure('Probe', []), 'Aあ');
console.log('[PASS] FSO Unicode TextStream round-trip (1 case)');
