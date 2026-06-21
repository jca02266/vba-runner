import { evalVBASingle, assert } from '../../test-libs/test-runner';

console.log("Running GetObject tests...");

const tests = [
    { code: "Debug.Print TypeName(GetObject(\"\", \"Scripting.Dictionary\"))", expected: "Dictionary" },
    // 括弧無し（引数無し）でも自動的に呼び出される必要がある（Nothing が返る）
    { code: "Debug.Print TypeName(GetObject)", expected: "Nothing" },
];

for (const t of tests) {
    let output = '';
    evalVBASingle(t.code, { onPrint: (o) => { output = o; } });
    assert.strictEqual(output.trim(), t.expected, t.code);
    console.log(`[PASS] ${t.code}`);
}

console.log('✅ GetObject: 全テスト通過');
