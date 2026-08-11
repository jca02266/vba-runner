import { strict as assert } from 'node:assert';
import { evalVBASingle } from '../../test-libs/test-runner';

const cases = [
    ['CInt("&H8000")', 'TYPE=Integer:VALUE=-32768'],
    ['CInt("&HFFFF")', 'TYPE=Integer:VALUE=-1'],
    ['CInt("&HFFFFFFFF")', 'TYPE=Integer:VALUE=-1'],
    ['CInt("&O37777777777")', 'TYPE=Integer:VALUE=-1'],
    ['CInt("&H7FFFFFFFFFFFFFFF")', 'ERR=13'],
    ['CLng("&H80000000")', 'TYPE=Long:VALUE=-2147483648'],
    ['CLng("&HFFFFFFFF")', 'TYPE=Long:VALUE=-1'],
    ['CLng("&H7FFFFFFFFFFFFFFF")', 'ERR=13'],
    ['CDec("&HFFFF")', 'TYPE=Decimal:VALUE=65535'],
    ['CDec("&H8000")', 'TYPE=Decimal:VALUE=32768'],
    ['CDec("&O177777")', 'TYPE=Decimal:VALUE=65535'],
    ['CCur("&H8000")', 'TYPE=Currency:VALUE=32768'],
    ['CCur("&HFFFF")', 'TYPE=Currency:VALUE=65535'],
    ['CCur("&O177777")', 'TYPE=Currency:VALUE=65535'],
    ['CCur("&HFFFFFFFFFFFFFFFF")', 'ERR=6'],
    ['CByte("&H7FFFFFFFFFFFFFFF")', 'ERR=13'],
    ['CLngLng("&O1777777777777777777777")', 'ERR=13'],
    ['Val("&HFFFF")', 'TYPE=Double:VALUE=-1'],
    ['TypeName(&H10000)', 'TYPE=String:VALUE=Long'],
    ['TypeName(&HFFFF&)', 'TYPE=String:VALUE=Long'],
    ['TypeName(&HFFFFFFFFFFFFFFFF^)', 'TYPE=String:VALUE=LongLong'],
    ['CLngLng("-&H8000000000000000")', 'TYPE=LongLong:VALUE=-9223372036854775808'],
] as const;

for (const [expression, expected] of cases) {
    const ev = evalVBASingle(String.raw`Option Explicit
Public Function Probe() As String
    Dim value As Variant, errNo As Long
    On Error Resume Next
    Err.Clear
    value = ${expression}
    errNo = Err.Number
    If errNo <> 0 Then
        Probe = "ERR=" & CStr(errNo)
    Else
        Probe = "TYPE=" & TypeName(value) & ":VALUE=" & CStr(value)
    End If
End Function`);
    assert.equal(ev.callProcedure('Probe', []), expected, `Radix sign/suffix ${expression}`);
}

assert.equal(cases.length, 22);
console.log(`[PASS] radix sign and suffix matrix (${cases.length} cases)`);
