import { evalVBASingle, assert } from '../../test-libs/test-runner';

console.log("Running CLngLng tests...");

const tests = [
    { code: "Debug.Print CLngLng(123.45)", expected: "123" },
    { code: "Debug.Print CLngLng(-123.45)", expected: "-123" },
    { code: "Debug.Print CLngLng(True)", expected: "-1" },
    { code: "Debug.Print CLngLng(\"9223372036854775807\")", expected: "9223372036854775807" },
    { code: "Debug.Print CLngLng(\"-9223372036854775808\")", expected: "-9223372036854775808" },
    { code: "Debug.Print TypeName(CLngLng(1))", expected: "LongLong" },
    { code: "Debug.Print VarType(CLngLng(1))", expected: "20" },
    { code: "Debug.Print CLngPtr(42)", expected: "42" },
    { code: "Debug.Print CLngPtr(-1)", expected: "-1" },
    { code: "Debug.Print VarType(CLngPtr(1))", expected: "20" },
    { code: "Sub T()\n    Dim x As LongPtr\n    Debug.Print x\nEnd Sub\nT", expected: "0" },
    { code: "Sub T()\n    Dim x As LongPtr\n    x = 100\n    Debug.Print x\nEnd Sub\nT", expected: "100" },
];

for (const t of tests) {
    let output = '';
    evalVBASingle(t.code, { onPrint: (o) => { output = o; } });
    assert.strictEqual(output.trim(), t.expected, t.code);
    console.log(`[PASS] ${t.code}`);
}

const rangeSource = String.raw`
Option Explicit
Public Function Probe() As String
    Dim value As LongLong, errNo As Long
    On Error Resume Next
    Err.Clear
    value = CLngLng("9223372036854775808")
    errNo = Err.Number
    Probe = "LLPOS=" & CStr(errNo)
    Err.Clear
    value = CLngLng("-9223372036854775809")
    Probe = Probe & " LLNEG=" & CStr(Err.Number)
    Err.Clear
    value = CLngPtr("9223372036854775808")
    Probe = Probe & " LPPOS=" & CStr(Err.Number)
    Err.Clear
    value = CLngLng("922x")
    Probe = Probe & " INVALID=" & CStr(Err.Number)
End Function
`;
const rangeEval = evalVBASingle(rangeSource);
assert.strictEqual(
    rangeEval.callProcedure('Probe', []),
    'LLPOS=6 LLNEG=6 LPPOS=6 INVALID=13',
    'CLngLng/CLngPtrは範囲外と構文不正を別エラーへ分類する',
);
console.log('[PASS] CLngLng/CLngPtr range error classification');

console.log('✅ CLngLng: 全テスト通過');
