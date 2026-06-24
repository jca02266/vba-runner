/**
 * VBARunner の `quiet` / `onPrint` オプションのテスト
 *
 * run() の [PASS] ログ抑制と Debug.Print の出力先を個別に制御できることを検証する。
 */
import { VBARunner, assert } from '../../test-libs/test-runner';

const DIR = 'tests/fixtures/multi-file';

// --- 1. quiet: false（デフォルト）では [PASS] ログが出力される ---
{
    const lines: string[] = [];
    const origLog = console.log;
    console.log = (s: string) => lines.push(s);
    try {
        const runner = new VBARunner(`${DIR}/MathUtils.bas`);
        runner.run('Add', [3, 4]);
    } finally {
        console.log = origLog;
    }
    assert.ok(lines.some(l => l.startsWith('[PASS] Add')), 'デフォルトでは [PASS] ログが出力される');
}
console.log('[PASS] quiet 未指定時に [PASS] ログが出る');

// --- 2. quiet: true では [PASS] ログが抑制される ---
{
    const lines: string[] = [];
    const origLog = console.log;
    console.log = (s: string) => lines.push(s);
    try {
        const runner = new VBARunner(`${DIR}/MathUtils.bas`, { quiet: true });
        const result = runner.run('Add', [3, 4]);
        assert.strictEqual(result, 7, '戻り値は quiet でも変わらない');
    } finally {
        console.log = origLog;
    }
    assert.ok(!lines.some(l => l.startsWith('[PASS]')), 'quiet: true で [PASS] ログが出ない');
}
console.log('[PASS] quiet: true で [PASS] ログが抑制される');

// --- 3. onPrint で Debug.Print の出力先を独立して切り替えられる ---
{
    const printed: string[] = [];
    const origLog = console.log;
    let consoleLogCalled = false;
    console.log = () => { consoleLogCalled = true; };
    try {
        const runner = new VBARunner(null, { onPrint: (s: string) => printed.push(s) });
        runner.eval('Debug.Print "hello"');
    } finally {
        console.log = origLog;
    }
    assert.deepStrictEqual(printed, ['hello'], 'onPrint に Debug.Print の出力が渡される');
    assert.ok(!consoleLogCalled, 'onPrint 指定時は console.log が呼ばれない');
}
console.log('[PASS] onPrint で Debug.Print の出力先を切り替えられる');

console.log('\n✅ VBARunner quiet/onPrint オプション: 全テスト通過');
