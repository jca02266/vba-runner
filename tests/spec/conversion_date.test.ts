import { evalVBASingle, assert } from '../../test-libs/test-runner';

console.log("Running CDate tests...");

const tests = [
    { code: "Debug.Print CDate(\"2023-05-11\")" },
    { code: "Debug.Print CDate(45057)" },
    { code: "Debug.Print CDate(\"45057\")" },
];

for (const t of tests) {
    let output = '';
    evalVBASingle(t.code, { onPrint: (o) => { output = o; } });
    assert.ok(
        output.includes("2023") && output.includes("05") && output.includes("11"),
        `${t.code} should include 2023, 05, 11 (got: ${output})`
    );
    console.log(`[PASS] ${t.code}`);
}

console.log('✅ CDate: 全テスト通過');
