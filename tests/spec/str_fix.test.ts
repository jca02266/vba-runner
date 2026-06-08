import { evalVBASingle, assert } from '../../test-libs/test-runner';

console.log("Running Str fixes tests...");

const tests = [
    { code: "Debug.Print Str(123)", expected: " 123" },
    { code: "Debug.Print Str(-123)", expected: "-123" },
    { code: "Debug.Print Str(0)", expected: " 0" },
    { code: "Debug.Print Str(Null)", expected: "Null" },
];

for (const t of tests) {
    let output = '';
    evalVBASingle(t.code, { onPrint: (o) => { output = o; } });
    // leading space は仕様なので trim しない
    assert.strictEqual(output, t.expected, t.code);
    console.log(`[PASS] ${t.code}`);
}

console.log('✅ Str Fixes: 全テスト通過');
