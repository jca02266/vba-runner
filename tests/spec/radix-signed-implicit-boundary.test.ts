import { strict as assert } from 'node:assert';
import { evalVBASingle } from '../../test-libs/test-runner';

const source = String.raw`
Option Explicit
Public Function ProbeSigned(ByVal kind As String, ByVal text As String) As String
    Dim source As Variant, result As Variant, errNo As Long
    On Error Resume Next
    Err.Clear
    source = CDec(text)
    Select Case kind
        Case "Byte": Dim byteValue As Byte: byteValue = source: result = byteValue
        Case "Integer": Dim integerValue As Integer: integerValue = source: result = integerValue
        Case "Long": Dim longValue As Long: longValue = source: result = longValue
        Case "LongLong": Dim longLongValue As LongLong: longLongValue = source: result = longLongValue
    End Select
    errNo = Err.Number
    If errNo <> 0 Then
        ProbeSigned = "ERR=" & CStr(errNo)
    Else
        ProbeSigned = "TYPE=" & TypeName(result) & ":VALUE=" & CStr(result)
    End If
    On Error GoTo 0
End Function

Public Function ProbeCurrency(ByVal text As String) As String
    Dim result As Currency, errNo As Long
    On Error Resume Next
    Err.Clear
    result = CCur(CDec(text))
    errNo = Err.Number
    If errNo <> 0 Then
        ProbeCurrency = "ERR=" & CStr(errNo)
    Else
        ProbeCurrency = "TYPE=" & TypeName(result) & ":VALUE=" & CStr(result)
    End If
    On Error GoTo 0
End Function
`;

const ev = evalVBASingle(source);
const cases = [
    ['Byte', '-1', 'ERR=6'],
    ['Integer', '-32769', 'ERR=6'],
    ['Integer', '-32768', 'TYPE=Integer:VALUE=-32768'],
    ['Long', '-2147483649', 'ERR=6'],
    ['Long', '-2147483648', 'TYPE=Long:VALUE=-2147483648'],
    ['LongLong', '-9223372036854775809', 'ERR=6'],
    ['LongLong', '-9223372036854775808', 'TYPE=LongLong:VALUE=-9223372036854775808'],
    ['LongLong', '9223372036854775807', 'TYPE=LongLong:VALUE=9223372036854775807'],
    ['LongLong', '9223372036854775808', 'ERR=6'],
];

for (const [kind, text, expected] of cases) {
  assert.equal(ev.callProcedure('ProbeSigned', [kind, text]), expected, `${kind} ${text}`);
}

assert.equal(ev.callProcedure('ProbeCurrency', ['922337203685477.5807']),
  'TYPE=Currency:VALUE=922337203685477.5807');
assert.equal(ev.callProcedure('ProbeCurrency', ['922337203685477.5808']), 'ERR=6');
assert.equal(ev.callProcedure('ProbeCurrency', ['900719925474099.3125']),
  'TYPE=Currency:VALUE=900719925474099.3125');

console.log(`[PASS] signed implicit conversion boundary matrix (${cases.length + 3} cases)`);
