import { evalVBASingle, assert } from '../../test-libs/test-runner';

console.log("Running CDec tests...");

const tests = [
    { code: "Debug.Print CDec(123.45)", expected: "123.45" },
    { code: "Debug.Print TypeName(CDec(1))", expected: "Decimal" },
    { code: "Debug.Print VarType(CDec(1))", expected: "14" },
];

for (const t of tests) {
    let output = '';
    evalVBASingle(t.code, { onPrint: (o) => { output = o; } });
    assert.strictEqual(output.trim(), t.expected, t.code);
    console.log(`[PASS] ${t.code}`);
}

console.log('✅ CDec: 全テスト通過');
