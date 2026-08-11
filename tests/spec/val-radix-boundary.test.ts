import { strict as assert } from 'node:assert';
import { evalVBASingle } from '../../test-libs/test-runner';

const source = String.raw`
Public Function Probe(ByVal text As String) As String
    On Error Resume Next
    Err.Clear
    Probe = CStr(Val(text)) & ":ERR=" & CStr(Err.Number)
    On Error GoTo 0
End Function
`;

const ev = evalVBASingle(source);
const cases: Array<[string, string]> = [
    ['&H7FFFFFFF', '2147483647:ERR=0'],
    ['&H80000000', '-2147483648:ERR=0'],
    ['&HFFFFFFFF', '-1:ERR=0'],
    ['&H7FFFFFFFFFFFFFFF', '9.22337203685478E+18:ERR=0'],
    ['&H8000000000000000', '-9.22337203685478E+18:ERR=0'],
    ['&HFFFFFFFFFFFFFFFF', '-1:ERR=0'],
];
for (const [input, expected] of cases) {
    assert.equal(ev.callProcedure('Probe', [input]), expected, input);
}

console.log(`[PASS] Val radix sign-extension boundary (${cases.length} cases)`);
