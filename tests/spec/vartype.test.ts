import { evalVBASingle, assert } from '../../test-libs/test-runner';

console.log("Running VarType tests...");

const tests = [
    { code: "Debug.Print VarType(Empty)", expected: 0 },
    { code: "Debug.Print VarType(Null)", expected: 1 },
    { code: "Debug.Print VarType(Nothing)", expected: 9 },
    { code: "Debug.Print VarType(True)", expected: 11 },
    { code: "Debug.Print VarType(#2023-01-01#)", expected: 7 },
    { code: "Debug.Print VarType(CVErr(5))", expected: 10 },
    { code: "Debug.Print VarType(1.23)", expected: 5 },
    { code: "Debug.Print VarType(\"hello\")", expected: 8 },
    { code: "Dim a(1): Debug.Print VarType(a)", expected: 8192 + 12 },
    { code: "Debug.Print vbInteger", expected: 2 },
    { code: "Debug.Print vbLong", expected: 3 },
    { code: "Debug.Print vbSingle", expected: 4 },
    { code: "Debug.Print vbDouble", expected: 5 },
    { code: "Debug.Print vbBoolean", expected: 11 },
    { code: "Debug.Print vbArray", expected: 8192 },
];

for (const t of tests) {
    let output = '';
    evalVBASingle(t.code, { onPrint: (o) => { output = o; } });
    assert.strictEqual(Number(output.trim()), t.expected, t.code);
    console.log(`[PASS] ${t.code}`);
}

console.log('✅ VarType: 全テスト通過');
