/**
 * VBARunner.run() で ByRef パラメーターを持つ Sub/Function を呼んだ際、
 * 呼び出し元の JS args 配列に最終値が書き戻されることのテスト
 *
 * 仕様バグ修正: callProcedure() は引数を localEnv にコピーするだけで、JS の
 * 数値・文字列・Boolean はプリミティブで参照を共有しないため、Sub 内で
 * ByRef パラメーターに代入（`n = n + 1` 等）しても、呼び出し元が渡した
 * args 配列には一切反映されなかった。VBA の既定の引数渡しは ByRef
 * （明示的な ByVal がない限り）であるため、これは多くの VBA コードが
 * 使う「ステータス + メッセージを ByRef で返す」パターンを壊していた。
 * callProcedure() の呼び出し後に、ByVal 以外のパラメーターの最終値を
 * args[i] へ書き戻すよう修正した。
 */
import { VBARunner, assert } from '../../test-libs/test-runner';

// --- 1. 単純な ByRef Long パラメーター ---
{
    const r = new VBARunner();
    r.eval('Sub Increment(ByRef n As Long)\n    n = n + 1\nEnd Sub');
    const args: any[] = [5];
    r.run('Increment', args);
    assert.strictEqual(args[0], 6, 'ByRef パラメーターへの代入が呼び出し元の args 配列に書き戻される');
}
console.log('[PASS] 単純な ByRef Long パラメーターが args 配列に書き戻される');

// --- 2. ByRef/ByVal が混在するパラメーター ---
{
    const r = new VBARunner();
    r.eval('Sub Mixed(ByRef a As Long, ByVal b As Long, ByRef msg As String)\n    a = a + 10\n    b = b + 100\n    msg = "done: " & a\nEnd Sub');
    const args: any[] = [1, 2, ''];
    r.run('Mixed', args);
    assert.strictEqual(args[0], 11, 'ByRef の a は書き戻される');
    assert.strictEqual(args[1], 2, 'ByVal の b は書き戻されない（呼び出し元の値が維持される）');
    assert.strictEqual(args[2], 'done: 11', 'ByRef の文字列 msg も書き戻される（status+messageパターン）');
}
console.log('[PASS] ByRef/ByVal混在パラメーターで ByRef のみ正しく書き戻される');

// --- 3. 既定（修飾子なし）は ByRef として書き戻される ---
{
    const r = new VBARunner();
    r.eval('Sub NoModifier(n As Long)\n    n = n * 2\nEnd Sub');
    const args: any[] = [3];
    r.run('NoModifier', args);
    assert.strictEqual(args[0], 6, '修飾子なしパラメーターは既定で ByRef として書き戻される');
}
console.log('[PASS] 修飾子なしパラメーターは既定で ByRef として書き戻される');

console.log('\n✅ VBARunner.run() の ByRef 書き戻し: 全テスト通過');
