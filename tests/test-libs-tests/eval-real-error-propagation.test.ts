/**
 * VBARunner.eval() の高速パス（単一式として一発評価する経路）が、式の実行中に発生した
 * 本物のランタイムエラーを握りつぶし、無関係な "Parse error: syntax error" に
 * 化けさせてしまうバグのテスト
 *
 * 仕様バグ修正: evalExpression() は「parseExpressionPublic() で単一式としてパースできるか」
 * を try/catch で試すが、その catch ブロックが parseExpressionPublic() の呼び出しだけでなく
 * 後続の実行（callProcedure/evaluateExpression）まで囲んでいたため、式の実行中に投げられた
 * 本物のランタイムエラー（Error 91 等）も「パース失敗」として握りつぶされ、文として再解析
 * した際に得られる無関係な構文エラー（または別の挙動）に置き換わっていた。
 * パース段階のみを try/catch で囲み、実行はその外側で行うよう修正した。
 */
import { VBARunner, assert } from '../../test-libs/test-runner';

// --- 1. 単一式の高速パスで実行された関数呼び出しが投げた本物のエラーが、
// そのまま eval() の呼び出し元に伝播する（run() と同じエラーになる） ---
{
    const r = new VBARunner();
    r.eval('Function Boom() As Long\n    Dim o As Object\n    o.DoSomething\nEnd Function');

    let runError = '';
    try { r.run('Boom', []); } catch (e: any) { runError = e.message; }

    let evalError = '';
    try { r.eval('Boom()'); } catch (e: any) { evalError = e.message; }

    assert.ok(/Run-time error '91'/.test(runError), 'run() は Error 91 を投げる');
    assert.strictEqual(evalError, runError, 'eval() の高速パスでも run() と同じ本物のエラーが伝播する（無関係な構文エラーに化けない）');
}
console.log('[PASS] eval() の高速パスは式実行中の本物のランタイムエラーをそのまま伝播する');

// --- 2. 既存の「パース失敗時はフォールバックする」挙動は維持される ---
{
    const r = new VBARunner();
    r.eval('Dim x : x = 10 : Debug.Print x');
    // "Dim x : x = 10 : Debug.Print x" は単一式としてパースできずフォールバックするが、
    // パースエラー自体は投げられない（正常にフォールバックして実行される）。
    const printed: string[] = [];
    const r2 = new VBARunner(null, { onPrint: (s: string) => printed.push(s) });
    r2.eval('Dim x : x = 10 : Debug.Print x');
    assert.deepStrictEqual(printed, ['10'], '単一式としてパースできない入力は引き続き文としてフォールバック実行される');
}
console.log('[PASS] 単一式としてパースできない入力のフォールバック実行は引き続き機能する');

console.log('\n✅ eval() 高速パスのエラー伝播: 全テスト通過');
