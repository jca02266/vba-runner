/**
 * Parse error メッセージへのソース行スニペット付与のテスト
 *
 * 行番号だけでは何が悪いか分かりにくいため、Parser に `sourceLines` を渡すと
 * 該当行のソース文字列がエラーメッセージに追記されることを検証する
 * （evalVBASingle/evalVBAModules は内部で自動的に渡す）。
 */
import { evalVBASingle, evalVBAModules, assert } from '../../test-libs/test-runner';

// --- 1. 単一モジュールの Parse error に該当行のソースが含まれる ---
{
    const code = [
        'Sub Test()',
        '    Const X As Integer 10',
        'End Sub',
    ].join('\n');

    let threw = false, msg = '';
    try {
        evalVBASingle(code);
    } catch (e: any) {
        threw = true;
        msg = e.message;
    }
    assert.ok(threw, 'Parse error が発生する');
    assert.ok(msg.includes('line 2'), `行番号が含まれる: ${msg}`);
    assert.ok(msg.includes('Const X As Integer 10'), `該当行のソースが含まれる: ${msg}`);
}
console.log('[PASS] 単一モジュールの Parse error にソース行が付与される');

// --- 2. 複数モジュールでも各モジュールの該当行ソースが付与される ---
{
    let threw = false, msg = '';
    try {
        evalVBAModules([
            { name: 'Mod1', code: 'Sub Ok()\nEnd Sub' },
            { name: 'Mod2', code: 'Sub Bad()\n    Dim 1abc\nEnd Sub' },
        ]);
    } catch (e: any) {
        threw = true;
        msg = e.message;
    }
    assert.ok(threw, 'Parse error が発生する');
    assert.ok(msg.includes('Dim 1abc'), `該当行のソースが含まれる: ${msg}`);
}
console.log('[PASS] 複数モジュールでも該当モジュールの行ソースが付与される');

console.log('\n✅ Parse error ソース行スニペット: 全テスト通過');
