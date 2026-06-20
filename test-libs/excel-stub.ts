/**
 * Excel API スタブモード
 *
 * `injectExcelStub(evaluator)` を呼ぶと、MockApplication が Tier 6 の
 * デフォルト束縛オブジェクトとして登録され、Excel 依存コードを追加の
 * モック設定なしで実行できる。
 *
 * 対応する VBA パターン（代表例）:
 *   Set ws = ActiveSheet
 *   ws.Cells(1, 1).Value = 100
 *   lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
 *   Application.ScreenUpdating = False
 *   Application.Calculation = xlCalculationManual
 *   Sheets("Data").Range("A1:B5").Value = arr
 *
 * すべての Excel API はノーオプスタブとして動作し、セルの読み書きは
 * インメモリで管理される。
 */
import { Evaluator } from '../src/engine/evaluator';
import { MockApplication } from '../src/engine/mock/MockExcel';

/**
 * evaluator に MockApplication を注入する。
 *
 * - `setDefaultBindingObject(app)` → Tier 6 で Range / Cells / ActiveSheet / Sheets 等を解決
 * - `setBuiltinOverride('Application', app)` → `Application.ScreenUpdating` 等のプロパティアクセスを有効化
 *
 * @param app 注入するインスタンス（省略時は `new MockApplication()`）。`MockApplication` を
 *   サブクラス化して `OnKey`/`OnTime` 等の未対応メンバーを追加したインスタンスを渡すこともできる。
 *   置き換えではなく拡張なので、`ActiveSheet`/`Sheets`/`Range` は変わらず動作し続ける。
 * @returns 注入した MockApplication（セルに事前データを設定したい場合に使う）
 */
export function injectExcelStub(evaluator: Evaluator, app: MockApplication = new MockApplication()): MockApplication {
    evaluator.setDefaultBindingObject(app);
    evaluator.setBuiltinOverride('Application', app);
    return app;
}
