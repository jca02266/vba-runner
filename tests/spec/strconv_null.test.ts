import { evalVBASingle, assert } from '../../test-libs/test-runner';

console.log("Running StrConv Null tests...");

const tests = [
    { code: "Debug.Print TypeName(StrConv(Null, 1))", expected: "Null" },
];

for (const t of tests) {
    let output = '';
    evalVBASingle(t.code, { onPrint: (o) => { output = o; } });
    assert.strictEqual(output.trim(), t.expected, t.code);
    console.log(`[PASS] ${t.code}`);
}

console.log('✅ StrConv Null: 全テスト通過');
