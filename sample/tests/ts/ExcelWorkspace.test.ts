import * as path from 'path';
import { VBARunner } from '../../../test-libs/test-runner';
import * as assert from 'assert';

const WORKSPACE = path.resolve('sample/excel-workspace');

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
    console.log('--- ExcelWorkspace サンプル検証 ---\n');

    // ─────────────────────────────────────────────────────────
    // RunWorkbookDemo: ワークブック/シートイベントのデモ
    // ─────────────────────────────────────────────────────────
    console.log('[Test] RunWorkbookDemo — エラーなく実行できる');
    const log1 = runWithCapture('RunWorkbookDemo');
    const out1 = log1.join('\n');

    assert.ok(log1.length > 0, 'RunWorkbookDemo: Debug.Print 出力が空でない');
    console.log(`[PASS] RunWorkbookDemo: ${log1.length} 行のログ出力`);

    const expectedWorkbook = [
        'Workbook_Open',        // Open イベント
        'ヘッダー初期化完了',    // Open 時の初期化
        'Sheet1 アクティブ',     // Activate イベント
        '自動計算',              // ThisWorkbook の m_ws1_Change による金額計算
        '変更検知',              // SheetChange バブリング
        'Workbook_BeforeClose',  // Close イベント
        '監査ログ',              // 監査ログ件数
        '総売上',                // 集計
    ];
    for (const kw of expectedWorkbook) {
        assert.ok(out1.includes(kw), `RunWorkbookDemo output contains "${kw}"`);
    }
    console.log('[PASS] RunWorkbookDemo: 期待するキーワードがすべて出力に含まれる');

    // 売上合計の検証: 緑茶10×120=1200, クッキー5×250=1250, コーヒー8×180=1440
    // 合計 = 3890
    assert.ok(out1.includes('3890'), `RunWorkbookDemo: 総売上 = 3890`);
    console.log('[PASS] RunWorkbookDemo: 総売上が正しい (¥3890)');

    // ─────────────────────────────────────────────────────────
    // RunDataDemo: データ操作・集計・レポートのデモ
    // ─────────────────────────────────────────────────────────
    console.log('\n[Test] RunDataDemo — エラーなく実行できる');
    const log2 = runWithCapture('RunDataDemo');
    const out2 = log2.join('\n');

    assert.ok(log2.length > 0, 'RunDataDemo: Debug.Print 出力が空でない');
    console.log(`[PASS] RunDataDemo: ${log2.length} 行のログ出力`);

    const expectedData = [
        'データ行数: 8',          // 8件のデータ
        '商品名順ソート',          // ソート実施
        '商品別集計',             // 集計ヘッダー
        '総売上',                 // 合計
        'VLookup',               // 検索
        'エラー捕捉',             // On Error Resume Next
        '売上レポート',            // レポート
    ];
    for (const kw of expectedData) {
        assert.ok(out2.includes(kw), `RunDataDemo output contains "${kw}"`);
    }
    console.log('[PASS] RunDataDemo: 期待するキーワードがすべて出力に含まれる');

    // 総売上検証:
    // 緑茶 10*120=1200, クッキー 5*250=1250, コーヒー 8*180=1440,
    // 緑茶 15*120=1800, シャンプー 3*680=2040, クッキー 12*250=3000,
    // チョコ 20*300=6000, コーヒー 6*180=1080
    // 合計 = 17,810
    assert.ok(out2.includes('17,810') || out2.includes('17810'), `RunDataDemo: 総売上 = 17,810`);
    console.log('[PASS] RunDataDemo: 総売上が正しい (¥17,810)');

    // エラーハンドリング確認
    assert.ok(out2.includes('91'), `RunDataDemo: Error 91 が正しく捕捉される`);
    console.log('[PASS] RunDataDemo: Error 91 が On Error Resume Next で捕捉される');

    // ─────────────────────────────────────────────────────────
    // RunValidationDemo: 入力検証イベントのデモ
    // ─────────────────────────────────────────────────────────
    console.log('\n[Test] RunValidationDemo — エラーなく実行できる');
    const log3 = runWithCapture('RunValidationDemo');
    const out3 = log3.join('\n');

    assert.ok(log3.length > 0, 'RunValidationDemo: Debug.Print 出力が空でない');
    console.log(`[PASS] RunValidationDemo: ${log3.length} 行のログ出力`);

    const expectedValidation = [
        '正常データ入力',
        '異常データ入力',
        '警告',              // 空商品名の警告
        'エラー',            // 負の数量・文字列単価
        '前後空白の自動除去', // 自動修正
        '検証エラー件数',    // エラー集計
        '日付検証',         // IsValidDate テスト
        '有効',             // 正常な日付
        '無効',             // 不正な日付
    ];
    for (const kw of expectedValidation) {
        assert.ok(out3.includes(kw), `RunValidationDemo output contains "${kw}"`);
    }
    console.log('[PASS] RunValidationDemo: 期待するキーワードがすべて出力に含まれる');

    // 検証エラー件数: 空商品名(1) + 負数量(1) + 文字列単価(1) = 3件
    assert.ok(out3.includes('検証エラー件数: 3'), `RunValidationDemo: エラー件数 = 3`);
    console.log('[PASS] RunValidationDemo: 検証エラー件数が正しい (3件)');

    console.log('\n✅ ExcelWorkspace: 全テスト通過');
}

main().catch(err => {
    console.error(err);
    process.exit(1);
});
