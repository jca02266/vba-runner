import { strict as assert } from 'node:assert';
import { evalVBASingle } from '../../test-libs/test-runner';

const source = String.raw`
Public Function Probe() As String
    Dim value As LongLong, commentValue As LongLong
    value = 123^
    commentValue = 123^   ' LongLong suffix before a comment
    Probe = CStr(2^-1) & "|" & CStr(10^-2) & "|" & CStr(2^+3) & "|" & _
        CStr(2^(-1)) & "|" & CStr(123^+3) & "|" & CStr(2^&H3) & "|" & _
        CStr(2^&O3) & "|" & CStr(value) & "|" & CStr(commentValue)
End Function
`;

const ev = evalVBASingle(source);
assert.equal(ev.callProcedure('Probe', []), '0.5|0.01|8|0.5|1860867|8|8|123|123');
console.log('[PASS] power operator keeps signed right operands distinct from LongLong suffixes');
