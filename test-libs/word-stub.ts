/**
 * Word API スタブモード
 *
 * `injectWordStub(evaluator)` を呼ぶと、ActiveDocument / Selection / Documents 等が
 * ノーオプスタブとして動き、Word 依存コードを追加の設定なしで実行できる。
 */
import { Evaluator } from '../src/engine/evaluator';
import { MockWordApplication } from '../src/engine/mock/MockWord';

/**
 * evaluator に MockWordApplication を注入する。
 *
 * - `setDefaultBindingObject(app)` → Tier 6 で ActiveDocument / Selection / Documents を解決
 * - `setBuiltinOverride('Application', app)` → Application.ScreenUpdating 等を有効化
 *
 * @returns 注入した MockWordApplication（doc.setContent() 等に使う）
 */
export function injectWordStub(evaluator: Evaluator): MockWordApplication {
    const app = new MockWordApplication();
    evaluator.setDefaultBindingObject(app);
    evaluator.setBuiltinOverride('Application', app);
    // CreateObject("Word.Application") / New Word.Application に対応
    evaluator.registerComObject(() => new MockWordApplication());
    return app;
}
