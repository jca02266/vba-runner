import { strict as assert } from 'node:assert';
import { evalVBASingle } from '../../test-libs/test-runner';

const source = String.raw`
Public Function Probe() As String
    Dim value As LongLong, commentValue As LongLong, concatValue As String
    value = 123^
    commentValue = 123^   ' LongLong suffix before a comment
    concatValue = CStr(123^) & "x"
    Probe = CStr(2 ^ -1) & "|" & CStr(10 ^ -2) & "|" & CStr(2 ^ +3) & "|" & _
        CStr(2 ^(-1)) & "|" & CStr(123 ^ +3) & "|" & CStr(2 ^ &H3) & "|" & _
        CStr(2 ^ &O3) & "|" & CStr(2^-1) & "|" & CStr(10^-2) & "|" & _
        CStr(value) & "|" & CStr(commentValue) & "|" & concatValue
End Function

Public Function UnaryProbe() As String
    UnaryProbe = CStr(-1) & "|" & CStr(- 1) & "|" & CStr(+1) & "|" & CStr(+ 1) & "|" & _
        CStr(1+-2) & "|" & CStr(1+ -2) & "|" & CStr(1 + -2) & "|" & CStr(1 + - 2) & "|" & _
        CStr(1--2) & "|" & CStr(1- -2) & "|" & CStr(1 - -2) & "|" & CStr(1 - - 2) & "|" & _
        CStr(2^-1) & "|" & CStr(2^ -1) & "|" & CStr(2 ^-1) & "|" & CStr(2 ^ -1) & "|" & _
        CStr(2^+1) & "|" & CStr(2^ +1) & "|" & CStr(2 ^+1) & "|" & CStr(2 ^ +1)
End Function
`;

const ev = evalVBASingle(source);
assert.equal(ev.callProcedure('Probe', []), '0.5|0.01|8|0.5|1860867|8|8|1|8|123|123|123x');
assert.equal(ev.callProcedure('UnaryProbe', []), '-1|-1|1|1|-1|-1|-1|-1|3|3|3|3|1|1|0.5|0.5|3|3|2|2');
console.log('[PASS] power operator keeps signed right operands distinct from LongLong suffixes');
