/**
 * Excel API スタブモード（excelStub: true）のテスト
 *
 * VBARunner に { excelStub: true } を渡すと、MockApplication が自動注入され、
 * Range / Cells / ActiveSheet / Application.ScreenUpdating 等が
 * 追加のモック設定なしで動作することを検証する。
 */
import { VBARunner, assert, MockApplication } from '../../test-libs/test-runner';

const suite = new VBARunner('tests/fixtures/excel-stub', { excelStub: true });

// --- Application プロパティの読み書きがエラーにならない ---
assert.strictEqual(suite.run('TestAppProperties', []), 'ok', 'Application プロパティのノーオプ');
console.log('[PASS] Application.ScreenUpdating / Calculation / EnableEvents');

// --- ActiveSheet.Cells().Value の読み書き ---
assert.strictEqual(suite.run('TestCellReadWrite', []), 42, 'Cells(1,1).Value = 42');
console.log('[PASS] ActiveSheet.Cells().Value 読み書き');

// --- ws.Rows.Count / ws.Columns.Count ---
const rc = suite.run('TestRowsColumnsCount', []);
assert.strictEqual(rc, 1048576 + 16384, 'Rows.Count + Columns.Count');
console.log('[PASS] ws.Rows.Count + ws.Columns.Count');

// --- End(xlUp).Row スタブ（データなし → 1 を返す） ---
assert.strictEqual(suite.run('TestEndXlUp', []), 1, 'End(xlUp).Row スタブ = 1');
console.log('[PASS] ws.Cells(lastRow, col).End(xlUp).Row');

// --- Sheets("name").Name ---
assert.strictEqual(suite.run('TestSheets', []), 'Data', 'Sheets("Data").Name = "Data"');
console.log('[PASS] Sheets("Data").Name');

// --- 各種ノーオプが例外を投げない ---
assert.strictEqual(suite.run('TestNoOps', []), 'ok', 'ノーオプ系呼び出しが例外なし');
console.log('[PASS] ノーオプ（Hidden / ColumnWidth / NumberFormat / Font.Bold / Interior.Color）');

// --- excelStub プロパティ経由で事前データを設定できる ---
const suite2 = new VBARunner('tests/fixtures/excel-stub', { excelStub: true });
suite2.excelStub!.ActiveSheet.setCellValue('A1', 99);
// ReadA1 は書き込まず A1 の値をそのまま返す
const val = suite2.run('ReadA1', []);
assert.strictEqual(val, 99, 'excelStub.ActiveSheet.setCellValue で事前データ設定');
console.log('[PASS] excelStub.ActiveSheet.setCellValue() で事前セルデータ');

// --- excelStub に MockApplication のサブクラスを渡すと、ActiveSheet 等を保ったまま拡張できる ---
class AppWithOnKey extends MockApplication {
    onKeyLog: string[] = [];
    OnKey(key: string, procName?: string) {
        this.onKeyLog.push(`${key}=${procName ?? ''}`);
    }
}
const customApp = new AppWithOnKey();
const suite3 = new VBARunner('tests/fixtures/excel-stub', { excelStub: customApp });
assert.strictEqual(suite3.excelStub, customApp, 'excelStub に渡したインスタンスがそのまま使われる');
// 拡張前から実装済みの ActiveSheet 系は変わらず動く
assert.strictEqual(suite3.run('TestCellReadWrite', []), 42, 'サブクラス使用時も Cells(1,1).Value = 42');
// 追加した OnKey も、標準の MockApplication が持つ ActiveSheet 系を壊さずに呼べる
suite3.run('TestCallOnKey', []);
assert.deepStrictEqual(customApp.onKeyLog, ['{LEFT}=SomeHandler'], 'サブクラスに追加した OnKey が呼ばれる');
console.log('[PASS] excelStub にカスタム MockApplication サブクラスを渡して拡張できる');

console.log('\n✅ excel-stub: 全テスト通過');
