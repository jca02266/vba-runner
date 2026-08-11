import { strict as assert } from 'node:assert';
import { evalVBASingle } from '../../test-libs/test-runner';

const ev = evalVBASingle(String.raw`
Function Probe(ByVal kind As String) As String
    Dim v As Variant, b As Byte, i As Integer, l As Long, ll As LongLong
    On Error Resume Next
    Err.Clear
    Select Case kind
        Case "Byte": v = "&H100": b = v
        Case "Integer": v = "&H10000": i = v
        Case "Long": v = "&H100000000": l = v
        Case "LongLong": v = "&H10000000000000000": ll = v
    End Select
    Probe = "ERR=" & CStr(Err.Number)
End Function
`);

const cases = [
    ['Byte', 'ERR=6'],
    ['Integer', 'ERR=6'],
    ['Long', 'ERR=6'],
    ['LongLong', 'ERR=6'],
] as const;

for (const [kind, expected] of cases) {
    assert.equal(ev.callProcedure('Probe', [kind]), expected,
        `implicit Variant-to-${kind} overflow boundary`);
}

assert.equal(cases.length, 4);
console.log(`[PASS] implicit radix target-width matrix (${cases.length} cases)`);
