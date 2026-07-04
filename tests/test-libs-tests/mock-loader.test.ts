/**
 * __mocks__ ディレクトリ注入のテスト
 *
 * テストパターン:
 *   1. __mocks__.bas — VBA モック（関数オーバーライド + VBA クラス）
 *   2. __mocks__.js  — JS モック（builtinOverrides + __addCreateObject__）
 *   3. __mocks__/ ディレクトリ形式（複数ファイルの混在）
 */
import { VBARunner, assert } from '../../test-libs/test-runner';

// --- 1. __mocks__.bas: VBA モックの関数とクラス ---
{
    const suite = new VBARunner('tests/fixtures/mocks-bas');

    // MsgBox が __mocks__.bas の実装（戻り値 42）に差し替わる
    const msgResult = suite.run('CallMsgBox', []);
    assert.strictEqual(msgResult, 42, '__mocks__.bas: MsgBox が差し替わり 42 を返す');
    console.log('[PASS] __mocks__.bas: MsgBox オーバーライド');

    // InputBox が __mocks__.bas の実装（"mock-input"）に差し替わる
    const inputResult = suite.run('CallInputBox', []);
    assert.strictEqual(inputResult, 'mock-input', '__mocks__.bas: InputBox が差し替わり "mock-input" を返す');
    console.log('[PASS] __mocks__.bas: InputBox オーバーライド');

    // ActiveSheet() が MockWorksheet インスタンスを返し、Name = "MockSheet"
    const sheetName = suite.run('UseSheet', []);
    assert.strictEqual(sheetName, 'MockSheet', '__mocks__.bas: VBA クラスモックが返す名前');
    console.log('[PASS] __mocks__.bas: VBA クラスモック (ActiveSheet / MockWorksheet)');
}

// --- 2. __mocks__.js: JS モック ---
{
    const suite = new VBARunner('tests/fixtures/mocks-js');

    // MsgBox が __mocks__.js の実装（戻り値 99）に差し替わる
    const msgResult = suite.run('CallMsgBox', []);
    assert.strictEqual(msgResult, 99, '__mocks__.js: MsgBox が差し替わり 99 を返す');
    console.log('[PASS] __mocks__.js: MsgBox JS オーバーライド');

    // VBScript.RegExp が __mocks__.js の VBScriptRegExp クラスで動く
    const reResult = suite.run('UseRegExp', []);
    assert.strictEqual(reResult, 'matched', '__mocks__.js: VBScript.RegExp のモックが動く');
    console.log('[PASS] __mocks__.js: __addCreateObject__ (VBScript.RegExp)');
}

// --- 3. __mocks__/ ディレクトリ形式 ---
{
    const suite = new VBARunner('tests/fixtures/mocks-dir');

    // __mocks__/MsgBox.js の MsgBox（戻り値 77）
    const msgResult = suite.run('GetMsgResult', []);
    assert.strictEqual(msgResult, 77, '__mocks__/ dir: MsgBox JS モック (77)');
    console.log('[PASS] __mocks__/ ディレクトリ: JS モック');

    // __mocks__/ExcelObjects.bas の Sheets() → MockWorksheet.Name = "MockSheet"
    const sheetName = suite.run('GetSheetName', []);
    assert.strictEqual(sheetName, 'MockSheet', '__mocks__/ dir: VBA クラスモック (Sheets → MockSheet)');
    console.log('[PASS] __mocks__/ ディレクトリ: VBA クラスモック');
}

// --- 4. ロード順・優先順位の確認 ---
// ソート順: Aardvark.js → Zebra.js → __mocks__.js（後勝ち: __mocks__.js が最高優先）
{
    const suite = new VBARunner('tests/fixtures/mocks-priority');

    // __mocks__.js が最後にロードされるため MsgBox = 99
    const result = suite.run('GetResult', []);
    assert.strictEqual(result, 99, '後勝ち: __mocks__.js が Aardvark.js / Zebra.js より優先');
    console.log('[PASS] 優先順位: __mocks__.js (99) > Zebra.js (22) > Aardvark.js (11)');
}

// --- 5. [回帰] 複数 VBA モックで同名関数の後勝ち (commit e39770c) ---
// promoteProcedures 修正前は同名関数が module-qualified キーのまま残り
// 修飾なし呼び出しで解決できなかった
{
    const suite = new VBARunner('tests/fixtures/mocks-vba-priority');

    // A_First.bas (11) < B_Second.bas (22): 後勝ちで B_Second が勝ち 22 を返す
    const result = suite.run('GetResult', []);
    assert.strictEqual(result, 22, '複数 VBA モックの同名関数: 後勝ち (B_Second=22 > A_First=11)');
    console.log('[PASS] 複数 VBA モック同名関数の後勝ち（promoteProcedures 回帰）');
}

console.log('\n✅ mock-loader: 全テスト通過');
