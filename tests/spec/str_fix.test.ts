import { evalVBASingle, assert } from '../../test-libs/test-runner';

console.log("Running Str fixes tests...");

const tests = [
    { code: "Debug.Print Str(123)", expected: " 123" },
    { code: "Debug.Print Str(-123)", expected: "-123" },
    { code: "Debug.Print Str(0)", expected: " 0" },
];

for (const t of tests) {
    let output = '';
    evalVBASingle(t.code, { onPrint: (o) => { output = o; } });
    // leading space は仕様なので trim しない
    assert.strictEqual(output, t.expected, t.code);
    console.log(`[PASS] ${t.code}`);
}

// Str(Null) は実 VBA と同様に Error 94 (Invalid use of Null) を発生させる
{
    let threw = false;
    try {
        evalVBASingle("Debug.Print Str(Null)");
    } catch(e: any) {
        threw = (e.number === 94);
    }
    assert.strictEqual(threw, true, 'Str(Null) → Error 94 (Invalid use of Null)');
    console.log('[PASS] Debug.Print Str(Null) → Error 94');
}

console.log('✅ Str Fixes: 全テスト通過');
