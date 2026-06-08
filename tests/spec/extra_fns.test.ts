import { evalVBASingle, assert } from '../../test-libs/test-runner';

console.log("Running Extra Functions tests...");

const tests = [
    { code: "Debug.Print CByte(12.5)", expected: "12" },
    { code: "Debug.Print CByte(13.5)", expected: "14" },
    { code: "Debug.Print Round(12.5)", expected: "12" },
    { code: "Debug.Print Round(13.5)", expected: "14" },
    { code: "Debug.Print StrComp(\"abc\", \"ABC\")", expected: "1" },
    { code: "Debug.Print StrComp(\"abc\", \"ABC\", 1)", expected: "0" },
    { code: "Debug.Print Format(1234.567, \"#,##0.00\")", expected: "1,234.57" },
    { code: "Debug.Print Format(0.123, \"Percent\")", expected: "12.30%" },
    { code: "Debug.Print Format(DateSerial(2023, 5, 11), \"yyyy-mm-dd\")", expected: "2023-05-11" },
];

for (const t of tests) {
    let output = '';
    evalVBASingle(t.code, { onPrint: (o) => { output = o; } });
    assert.strictEqual(output.trim(), t.expected, t.code);
    console.log(`[PASS] ${t.code}`);
}

console.log('✅ Extra Functions: 全テスト通過');
