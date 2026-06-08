import { evalVBASingle, assert } from '../../test-libs/test-runner';

console.log("Running GetObject tests...");

const tests = [
    { code: "Debug.Print TypeName(GetObject(\"\", \"Scripting.Dictionary\"))", expected: "Dictionary" },
];

for (const t of tests) {
    let output = '';
    evalVBASingle(t.code, { onPrint: (o) => { output = o; } });
    assert.strictEqual(output.trim(), t.expected, t.code);
    console.log(`[PASS] ${t.code}`);
}

console.log('✅ GetObject: 全テスト通過');
