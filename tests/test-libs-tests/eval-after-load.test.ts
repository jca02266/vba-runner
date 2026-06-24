/**
 * VBARunner がファイルをロードした後に eval() で Dim ... As New を使うケースのテスト
 *
 * ファイルロード中に呼ばれる evaluator.setSourceModule() の影響が construction 後も
 * 残るため、eval() で module-level の Dim（Public/Friend 以外）を書くと、書き込み先
 * （ロード時に残った currentSourceModule のモジュール専用スコープ）と読み込み元
 * （グローバルスコープ）がズレて変数が見つからなくなる不具合があった。
 * eval() は常にどのファイルにも属さない専用のトップレベルとして評価されるべき。
 */
import { VBARunner, assert } from '../../test-libs/test-runner';

const DIR = 'tests/fixtures/eval-after-load';

// --- 1. ディレクトリロード後、eval を2回に分けて New インスタンスのプロパティを読む ---
{
    const runner = new VBARunner(DIR);
    runner.eval('Dim t As New Task');
    const title = runner.eval('t.Title');
    assert.strictEqual(title, '', 'New 直後のプロパティはデフォルト値（空文字）');
}
console.log('[PASS] ディレクトリロード後、eval2回でNewインスタンスのプロパティを読める');

// --- 2. 単一ファイル指定でロードした場合も同様（クラス定義も同一ファイル内） ---
{
    const runner = new VBARunner(`${DIR}/standalone/StandaloneTask.bas`);
    runner.eval('Dim t As New Task');
    const title = runner.eval('t.Title');
    assert.strictEqual(title, '', '単一ファイル指定でロードしても同様に動作する');
}
console.log('[PASS] 単一ファイル指定でロード後も同様に動作する');

// --- 3. colon 区切りで Dim → 代入 → Debug.Print まで一括 eval ---
{
    const printed: string[] = [];
    const runner = new VBARunner(DIR, { onPrint: (s: string) => printed.push(s) });
    runner.eval('Dim t As New Task: t.Title = "hello": Debug.Print t.Title');
    assert.deepStrictEqual(printed, ['hello'], 'colon連結でNew→代入→読み込みが一貫して動く');
}
console.log('[PASS] colon連結でNew→代入→Debug.Printが一貫して動く');

console.log('\n✅ eval() after file load: 全テスト通過');
