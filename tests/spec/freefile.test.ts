import { evalVBASingle, assert } from '../../test-libs/test-runner';

console.log("Running FreeFile tests...");

const tests = [
    { code: "Debug.Print FreeFile()", expected: "1" },
    { code: "Debug.Print FreeFile(1)", expected: "256" },
    // 括弧無し（引数無し）でも自動的に呼び出される必要がある
    { code: "Debug.Print FreeFile", expected: "1" },
];

for (const t of tests) {
    let output = '';
    evalVBASingle(t.code, { onPrint: (o) => { output = o; } });
    assert.strictEqual(output.trim(), t.expected, t.code);
    console.log(`[PASS] ${t.code}`);
}

for (const code of ['Debug.Print FreeFile(2)', 'Debug.Print FreeFile(1.5)']) {
    let error: any = null;
    try { evalVBASingle(code); } catch (e: any) { error = e; }
    assert.strictEqual(error?.number, 5, `${code} はError 5`);
    console.log(`[PASS] ${code} -> Error 5`);
}

{
    let error: any = null;
    try { evalVBASingle('Debug.Print FreeFile(Null)'); } catch (e: any) { error = e; }
    assert.strictEqual(error?.number, 94, 'FreeFile(Null) はError 94');
    console.log('[PASS] FreeFile(Null) -> Error 94');
}

console.log('✅ FreeFile: 全テスト通過');
