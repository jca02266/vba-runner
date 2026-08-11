import { strict as assert } from 'node:assert';
import { evalVBASingle } from '../../test-libs/test-runner';

const ev = evalVBASingle(String.raw`
Option Explicit
Public Function ProbeConversion(ByVal kind As String, ByVal text As String) As String
    Dim result As Variant, errNo As Long
    On Error Resume Next
    Err.Clear
    Select Case kind
        Case "CByte": result = CByte(text)
        Case "CInt": result = CInt(text)
        Case "CLng": result = CLng(text)
        Case "CLngLng": result = CLngLng(text)
        Case "CDec": result = CDec(text)
        Case "CCur": result = CCur(text)
        Case "Val": result = Val(text)
    End Select
    errNo = Err.Number
    If errNo <> 0 Then
        ProbeConversion = "ERR=" & CStr(errNo)
    Else
        ProbeConversion = "TYPE=" & TypeName(result) & ":VALUE=" & CStr(result)
    End If
End Function
`);

const values = [
    '0', '1', '-1', '&H7FFF', '&H8000', '&HFFFF',
    '&H7FFFFFFF', '&H80000000', '&HFFFFFFFF',
    '&H7FFFFFFFFFFFFFFF', '&H8000000000000000', '&HFFFFFFFFFFFFFFFF',
    '&O177777', '&O37777777777', '&O1777777777777777777777',
    '9007199254740993', '-9223372036854775808',
];
const kinds = ['CByte', 'CInt', 'CLng', 'CLngLng', 'CDec', 'CCur', 'Val'];
const matrix = kinds.flatMap(kind => values.map(value => ({
    kind, value, result: ev.callProcedure('ProbeConversion', [kind, value]),
})));
assert.equal(matrix.length, kinds.length * values.length);
assert.equal(matrix.find(x => x.kind === 'CInt' && x.value === '&H7FFF')?.result,
    'TYPE=Integer:VALUE=32767');
assert.equal(matrix.find(x => x.kind === 'CLng' && x.value === '&H7FFFFFFF')?.result,
    'TYPE=Long:VALUE=2147483647');
assert.equal(matrix.find(x => x.kind === 'Val' && x.value === '&HFFFF')?.result,
    'TYPE=Double:VALUE=-1');

console.log(`[PASS] declaration and conversion boundary matrix (${matrix.length} cases)`);
