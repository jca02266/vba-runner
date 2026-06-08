import { evalVBASingle, assert } from '../../test-libs/test-runner';

console.log("Running FreeFile tests...");

const tests = [
    { code: "Debug.Print FreeFile()", expected: "1" },
    { code: "Debug.Print FreeFile(1)", expected: "256" },
];

for (const t of tests) {
    let output = '';
    evalVBASingle(t.code, { onPrint: (o) => { output = o; } });
    assert.strictEqual(output.trim(), t.expected, t.code);
    console.log(`[PASS] ${t.code}`);
}

console.log('✅ FreeFile: 全テスト通過');
