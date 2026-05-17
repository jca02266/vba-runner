/**
 * CurDir / CurDir$ 関数のテスト (§6.1.2.5.1.1)
 *
 * VBA のカレントディレクトリ取得関数。
 * VBA Runner では Sandbox ルートを Windows 形式の仮想パスとして返す。
 */
import { Lexer } from '../../src/engine/lexer';
import { Parser } from '../../src/engine/parser';
import { Evaluator } from '../../src/engine/evaluator';
import { assert } from '../../test-libs/test-runner';

console.log('--- Starting CurDir Tests ---');

// --- 1. デフォルト Sandbox での CurDir ---
{
    const tokens = new Lexer('').tokenize();
    const ast = new Parser(tokens).parse();
    const ev = new Evaluator(console.log);
    ev.evaluate(ast);

    // デフォルト sandbox ルート (/sandbox) → 仮想パスは "\"
    assert.strictEqual(ev.evalExpression('CurDir()'), '\\', 'CurDir() はデフォルトで Sandbox ルート');
    assert.strictEqual(ev.evalExpression('CurDir$()'), '\\', 'CurDir$() も同じ動作');

    // ドライブレター指定（モック実装なのでサンドボックスルートを返す）
    assert.strictEqual(ev.evalExpression('CurDir("C")'), '\\', 'CurDir("C") も同じ動作');
    console.log('[PASS] デフォルト Sandbox での CurDir');
}

// --- 2. カスタム Sandbox ルートでの CurDir ---
{
    const tokens = new Lexer('').tokenize();
    const ast = new Parser(tokens).parse();
    // sandboxRoot を "/custom/root" に変更
    const ev = new Evaluator(console.log, { sandboxRoot: '/custom/root' });
    ev.evaluate(ast);

    // カスタムルートでも、ルート自体は仮想パスでは "\"
    assert.strictEqual(ev.evalExpression('CurDir()'), '\\', 'カスタム Sandbox ルートでも CurDir() = "\\"');
    console.log('[PASS] カスタム Sandbox ルートでの CurDir');
}

console.log('\n✅ CurDir: 全テスト通過');
