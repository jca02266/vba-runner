import { strict as assert } from 'node:assert';
import { evalVBASingle } from '../../test-libs/test-runner';

const cases = [
    ['0^', '0|LongLong|20'],
    ['9223372036854775807^', '9223372036854775807|LongLong|20'],
    ['-9223372036854775808^', '-9223372036854775808|LongLong|20'],
    ['&H20000000000001^', '9007199254740993|LongLong|20'],
    ['&HFFFFFFFFFFFFFFFF^', '-1|LongLong|20'],
    ['&O400000000000000000001^', '4611686018427387905|LongLong|20'],
] as const;

for (const [literal, expected] of cases) {
    const ev = evalVBASingle(String.raw`Option Explicit
Const K As LongLong = ${literal}
Public Function Probe() As String
    Probe = CStr(K) & "|" & TypeName(K) & "|" & CStr(VarType(K))
End Function`);
    assert.equal(ev.callProcedure('Probe', []), expected, `LongLong Const ${literal}`);
}

assert.equal(cases.length, 6);
console.log(`[PASS] LongLong constant inference matrix (${cases.length} cases)`);
