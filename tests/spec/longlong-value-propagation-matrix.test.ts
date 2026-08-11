import { strict as assert } from 'node:assert';
import { evalVBASingle } from '../../test-libs/test-runner';

const cases = [
    ['&H7FFFFFFFFFFFFFFF^', '9223372036854775807'],
    ['&H8000000000000000^', '-9223372036854775808'],
    ['&H20000000000001^', '9007199254740993'],
    ['&O400000000000000000001^', '4611686018427387905'],
] as const;

for (const [literal, expected] of cases) {
    const ev = evalVBASingle(String.raw`Option Explicit
Function ReturnLongLong() As LongLong
    ReturnLongLong = ${literal}
End Function
Public Function Probe() As String
    Dim value As LongLong, variantValue As Variant, values(0 To 2) As Variant
    value = ${literal}
    variantValue = value
    values(0) = value
    values(1) = ${literal}
    values(2) = ReturnLongLong()
    Probe = TypeName(variantValue) & ":" & CStr(variantValue) & "|" & _
        TypeName(values(0)) & ":" & CStr(values(0)) & "|" & _
        TypeName(values(1)) & ":" & CStr(values(1)) & "|" & _
        TypeName(values(2)) & ":" & CStr(values(2)) & "|" & CStr(VarType(variantValue))
End Function`);
    assert.equal(ev.callProcedure('Probe', []),
        `LongLong:${expected}|LongLong:${expected}|LongLong:${expected}|LongLong:${expected}|20`,
        `LongLong propagation ${literal}`);
}

assert.equal(cases.length, 4);
console.log(`[PASS] LongLong value propagation matrix (${cases.length} cases)`);
