import { strict as assert } from 'node:assert';
import { evalVBASingle } from '../../test-libs/test-runner';

const ev = evalVBASingle(String.raw`Option Explicit
Public Function Probe() As String
    Probe = CStr(CDbl(#1/1/100#)) & "|" & _
        CStr(CDbl(#1/1/101#)) & "|" & _
        CStr(CDbl(#3/15/2024 12:34:56#))
End Function`);

const actual = ev.callProcedure('Probe', []);
assert.equal(actual, '-657434|-657069|45366.52425925926',
    'date literals must use calendar values rather than host timezone offsets');
console.log(`[PASS] Date literal calendar serials are timezone-independent: ${actual}`);
