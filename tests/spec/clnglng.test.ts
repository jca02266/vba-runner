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

console.log('✅ CLngLng: 全テスト通過');
