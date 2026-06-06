/**
 * Access API スタブモード
 *
 * `injectAccessStub(evaluator)` を呼ぶと、CurrentDb / DoCmd / Forms 等が
 * ノーオプスタブとして動き、Access 依存コードを追加の設定なしで実行できる。
 *
 * 対応する VBA パターン（代表例）:
 *   Set db = CurrentDb()
 *   Set rs = db.OpenRecordset("TableName")
 *   Do While Not rs.EOF
 *       Debug.Print rs.Fields("Name").Value
 *       rs.MoveNext
 *   Loop
 *   DoCmd.SetWarnings False
 *   Forms("frmCustomer").txtName.Value
 *
 * フォームコントロールのアクセス（Me.txtName 等）は事前に
 * `app.Forms("frmName").setControl("txtName", value)` で登録が必要。
 */
import { Evaluator } from '../src/engine/evaluator';
import { MockAccessApplication } from '../src/engine/mock/MockAccess';

/**
 * evaluator に MockAccessApplication を注入する。
 *
 * - `setDefaultBindingObject(app)` → Tier 6 で CurrentDb / DoCmd / Forms 等を解決
 * - `setBuiltinOverride('Application', app)` → Application.Echo 等を有効化
 * - `setBuiltinOverride('CurrentDb', ...)` → CurrentDb() を直接呼べるようにする
 * - `setBuiltinOverride('DoCmd', ...)` → DoCmd.OpenForm 等を有効化
 *
 * @returns 注入した MockAccessApplication（db.setTableData() 等に使う）
 */
export function injectAccessStub(evaluator: Evaluator): MockAccessApplication {
    const app = new MockAccessApplication();
    evaluator.setDefaultBindingObject(app);
    evaluator.setBuiltinOverride('Application', app);
    evaluator.setBuiltinOverride('CurrentDb', () => app.CurrentDb());
    evaluator.setBuiltinOverride('DoCmd', app.DoCmd);
    evaluator.setBuiltinOverride('Forms', (name: any) => app.Forms(name));
    return app;
}
