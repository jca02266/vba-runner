/**
 * 未定義プロシージャ呼び出し時のエラー表示テスト
 *
 * 修正前:
 * 1. `End Su` のようなタイポは ParseError を生成せず、Sub ボディが EOF まで
 *    飲み込まれてトップレベルの呼び出しが実行されない（無音で終了）
 * 2. `Mainloo` のようなタイポはランタイムエラーを生成せず 0 を返す
 *    （未宣言変数の暗黙初期化と区別できなかった）
 */
import { Lexer } from '../../src/engine/lexer';
import { Parser } from '../../src/engine/parser';
import { Evaluator, VbaErrorCode } from '../../src/engine/evaluator';
import { assert } from '../../test-libs/test-runner';

// --- 1. End Su (End Sub のタイポ) で Parse Error が生成される ---
{
    const src = `
Sub MainLoop()
  Debug.Print "hello"
End Su

MainLoop
`;
    const tokens = new Lexer(src).tokenize();
    const ast = new Parser(tokens, { errorRecovery: true }).parse();
    assert.strictEqual(
        (ast.diagnostics?.length ?? 0) > 0, true,
        'End Su は Parse Error を diagnostics に追加する'
    );
    const msg = ast.diagnostics[0].message;
    assert.strictEqual(
        msg.includes("End") || msg.includes("unexpected"), true,
        `診断メッセージに 'End' または 'unexpected' が含まれる (got: ${msg})`
    );
    console.log("[PASS] End Su → Parse Error diagnostic 生成");
}

// --- 2. 未定義の Sub を呼び出すとランタイムエラーになる ---
{
    const src = `
Sub MainLoop()
  Debug.Print "hello"
End Sub

Mainloo
`;
    let caughtError: any = null;
    try {
        const ast = new Parser(new Lexer(src).tokenize()).parse();
        const ev = new Evaluator(console.log);
        ev.evaluateModule(ast);
        ev.resolveIdentifiers([{ ast, moduleName: '' }]);
    } catch (e: any) {
        caughtError = e;
    }
    assert.strictEqual(caughtError !== null, true, 'Mainloo 呼び出しでエラーが発生する');
    assert.strictEqual(
        caughtError?.number === VbaErrorCode.SUB_OR_FUNCTION_NOT_DEFINED, true,
        `エラーコードが SUB_OR_FUNCTION_NOT_DEFINED (35) (got: ${caughtError?.number})`
    );
    assert.strictEqual(
        String(caughtError?.message ?? '').includes('Mainloo'), true,
        `エラーメッセージに 'Mainloo' が含まれる (got: ${caughtError?.message})`
    );
    console.log("[PASS] 未定義 Sub 呼び出し → SUB_OR_FUNCTION_NOT_DEFINED エラー");
}

// --- 3. 正しい名前で呼び出すと正常に動作する ---
{
    const src = `
Sub MainLoop()
  Debug.Print "hello"
End Sub

MainLoop
`;
    let output = '';
    const ast = new Parser(new Lexer(src).tokenize()).parse();
    const ev = new Evaluator((s) => { output += s + '\n'; });
    ev.evaluateModule(ast);
    ev.resolveIdentifiers([{ ast, moduleName: '' }]);
    assert.strictEqual(output.trim(), 'hello', '正しい名前で呼び出すと正常に実行される');
    console.log("[PASS] 正しい Sub 名で呼び出し → 正常実行");
}

console.log("✅ All undefined-proc-error tests passed");
