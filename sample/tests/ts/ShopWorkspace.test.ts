import * as path from 'path';
import { VBARunner } from '../../../test-libs/test-runner';
import * as assert from 'assert';

const WORKSPACE = path.resolve('sample/workspace');

/** VBARunner 生成前に console.log を差し替えて Debug.Print 出力をキャプチャする */
function runWithCapture(proc: string): string[] {
    const captured: string[] = [];
    const orig = console.log;
    console.log = (...args: any[]) => captured.push(args.map(String).join(' '));
    const runner = new VBARunner(WORKSPACE);
    try {
        runner.eval(proc);
    } finally {
        console.log = orig;
    }
    return captured;
}

async function main() {
    console.log('--- ShopWorkspace サンプル検証 ---\n');

    // ─────────────────────────────────────────────────────────
    // RunDemo: クラスを直接使うデモ
    // ─────────────────────────────────────────────────────────
    console.log('[Test] RunDemo — エラーなく実行できる');
    const log1 = runWithCapture('RunDemo');
    const output1 = log1.join('\n');

    assert.ok(log1.length > 0, 'RunDemo: Debug.Print output should not be empty');
    console.log(`[PASS] RunDemo: ${log1.length} 行のログ出力`);

    const expectedRunDemo = [
        'SimpleShop',
        '在庫一覧',
        '緑茶',
        'クッキー',
        'コーヒー',
        'シャンプー',
        'チョコ',
        '在庫評価額合計',
        '在庫不足',    // TrySell(inv, 1, 99) のエラーハンドリング
        '価格変更',    // AdjustPrice が成功した証拠
        '在庫補充',    // AdjustStock が成功した証拠
    ];
    for (const kw of expectedRunDemo) {
        assert.ok(output1.includes(kw), `RunDemo output contains "${kw}"`);
    }
    console.log('[PASS] RunDemo: 期待するキーワードがすべて出力に含まれる');

    // TotalValue の検証
    // 初期: 緑茶120×10, クッキー250×20, コーヒー180×5, シャンプー680×15, チョコ300×8
    // TrySell(1,8): 緑茶 10-8=2  / TrySell(2,3): クッキー 20-3=17
    // AdjustPrice(1,150): 緑茶 → ¥150  / AdjustStock(1,20): 緑茶 2+20=22
    // TotalValue = 150*22 + 250*17 + 180*5 + 680*15 + 300*8 = 21050
    assert.ok(output1.includes('21050'), `RunDemo: 在庫評価額合計 = 21050`);
    console.log('[PASS] RunDemo: 在庫評価額合計が正しい (¥21050)');

    // ─────────────────────────────────────────────────────────
    // RunFormDemo: ShopForm（UserForm 相当）を経由するデモ
    // ─────────────────────────────────────────────────────────
    console.log('\n[Test] RunFormDemo — エラーなく実行できる');
    const log2 = runWithCapture('RunFormDemo');
    const output2 = log2.join('\n');

    assert.ok(log2.length > 0, 'RunFormDemo: Debug.Print output should not be empty');
    console.log(`[PASS] RunFormDemo: ${log2.length} 行のログ出力`);

    const expectedRunFormDemo = [
        '在庫管理フォーム',
        'データ読み込み完了',
        '販売完了',
        '補充完了',
        '価格変更',
        '在庫一覧',
    ];
    for (const kw of expectedRunFormDemo) {
        assert.ok(output2.includes(kw), `RunFormDemo output contains "${kw}"`);
    }
    console.log('[PASS] RunFormDemo: 期待するキーワードがすべて出力に含まれる');

    console.log('\n✅ ShopWorkspace: 全テスト通過');
}

main().catch(err => {
    console.error(err);
    process.exit(1);
});
