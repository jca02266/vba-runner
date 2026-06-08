import { evalVBASingle, assert } from '../../test-libs/test-runner';

console.log("Running Environ tests...");

const tests = [
    { code: "Debug.Print Environ(\"VBA_TEST_VAR\")", expected: "HelloWorld" },
    { code: "Debug.Print Environ(\"NON_EXISTENT_VAR\")", expected: "" },
    { code: "Debug.Print Environ$(1)", expected: "VBA_TEST_VAR=HelloWorld" },
    { code: "Debug.Print Environ(9999)", expected: "" },
];

const envOpts = { env: { VBA_TEST_VAR: "HelloWorld" } };

for (const t of tests) {
    let output = '';
    evalVBASingle(t.code, { ...envOpts, onPrint: (o) => { output = o; } });
    assert.strictEqual(output.trim(), t.expected, t.code);
    console.log(`[PASS] ${t.code}`);
}

console.log('✅ Environ: 全テスト通過');
