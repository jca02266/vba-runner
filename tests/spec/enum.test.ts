/**
 * Enum 宣言のテスト (§5.2.3.4)
 *
 * 検証する挙動:
 *   - 明示的な値の指定
 *   - 値を省略した場合の自動インクリメント
 *   - メンバ名による直接アクセス（Red 等）
 *   - Enum 名経由のアクセス（Color.Red 等）
 */
import { Lexer } from '../../src/engine/lexer';
import { Parser } from '../../src/engine/parser';
import { Evaluator } from '../../src/engine/evaluator';
import { assert } from '../../test-libs/test-runner';

function evalVBA(code: string): any {
    const tokens = new Lexer(code).tokenize();
    const ast = new Parser(tokens).parse();
    const ev = new Evaluator(console.log);
    ev.evaluate(ast);
    return ev;
}

console.log('--- Starting Enum Declaration Tests ---');

const enumCode = `
    Enum Color
        Red = 1
        Green
        Blue = 10
        Yellow
    End Enum
`;

const ev = evalVBA(enumCode);

// --- メンバ名による直接アクセス ---
{
    assert.strictEqual(ev.evalExpression('Red'), 1, 'Red = 1 (明示)');
    assert.strictEqual(ev.evalExpression('Green'), 2, 'Green = 2 (自動インクリメント)');
    assert.strictEqual(ev.evalExpression('Blue'), 10, 'Blue = 10 (明示)');
    assert.strictEqual(ev.evalExpression('Yellow'), 11, 'Yellow = 11 (自動インクリメント)');
    console.log('[PASS] メンバ名による直接アクセス');
}

// --- Enum 名経由のアクセス ---
{
    assert.strictEqual(ev.evalExpression('Color.Red'), 1, 'Color.Red = 1');
    assert.strictEqual(ev.evalExpression('Color.Green'), 2, 'Color.Green = 2');
    assert.strictEqual(ev.evalExpression('Color.Blue'), 10, 'Color.Blue = 10');
    assert.strictEqual(ev.evalExpression('Color.Yellow'), 11, 'Color.Yellow = 11');
    console.log('[PASS] Enum 名経由のアクセス');
}

console.log('\n✅ Enum 宣言: 全テスト通過');
