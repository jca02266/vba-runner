/**
 * Evaluator の DebugHook ユニットテスト
 *
 * Worker スレッドを使わず Evaluator に直接 DebugHook を設定して
 * フック呼び出しのタイミング・引数を検証する。
 */
import { Lexer } from '../../src/engine/lexer';
import { Parser } from '../../src/engine/parser';
import { Evaluator, Environment, DebugHook } from '../../src/engine/evaluator';
import { assert } from '../../test-libs/test-runner';

function makeEvaluator(src: string, hook?: DebugHook): {
    evaluator: Evaluator;
    lines: number[];
    depths: number[];
    envs: Environment[];
} {
    const lines: number[] = [];
    const depths: number[] = [];
    const envs: Environment[] = [];

    const captureHook: DebugHook = {
        onBeforeStatement(line, depth, env) {
            lines.push(line);
            depths.push(depth);
            envs.push(env);
            hook?.onBeforeStatement(line, depth, env, []);
        },
    };

    const ast = new Parser(new Lexer(src).tokenize()).parse();
    const evaluator = new Evaluator(() => { /* suppress output */ });
    evaluator.setDebugHook(captureHook);
    evaluator.setSourceModule('Module1');
    evaluator.evaluateModule(ast);
    evaluator.resolveIdentifiers([{ ast, moduleName: 'Module1' }]);

    return { evaluator, lines, depths, envs };
}

// 1. モジュールロード中はフックが発火しない
{
    const src = 'Sub Main()\n  x = 1\nEnd Sub';
    const { lines } = makeEvaluator(src);
    // evaluateModule だけでは callProcedure を呼ばないのでフック不発
    assert.strictEqual(lines.length, 0, 'module load does not fire hook');
    console.log('[PASS] Hook silent during module load');
}

// 2. callProcedure でフックが発火する
{
    const src = 'Sub Main()\n  Dim x As Long\n  x = 1\nEnd Sub';
    const { evaluator, lines } = makeEvaluator(src);
    evaluator.callProcedure('Main', []);
    assert.ok(lines.length > 0, 'hook fires during procedure execution');
    console.log('[PASS] Hook fires during callProcedure, lines:', lines);
}

// 3. フック発火行番号が正しい（ステートメントの行）
{
    const src = [
        'Sub Main()',   // line 1
        '  x = 1',     // line 2
        '  x = 2',     // line 3
        '  x = 3',     // line 4
        'End Sub',     // line 5
    ].join('\n');
    const { evaluator, lines } = makeEvaluator(src);
    evaluator.callProcedure('Main', []);
    assert.ok(lines.includes(2), 'line 2 fired');
    assert.ok(lines.includes(3), 'line 3 fired');
    assert.ok(lines.includes(4), 'line 4 fired');
    assert.ok(!lines.includes(1), 'Sub declaration line not fired');
    assert.ok(!lines.includes(5), 'End Sub line not fired');
    console.log('[PASS] Hook fires at correct line numbers:', lines);
}

// 4. コール深度がプロシージャネスト数を反映する
{
    const src = [
        'Sub Outer()',  // line 1
        '  Call Inner()',  // line 2
        'End Sub',      // line 3
        'Sub Inner()',  // line 4
        '  x = 99',    // line 5
        'End Sub',     // line 6
    ].join('\n');
    const { evaluator, lines, depths } = makeEvaluator(src);
    evaluator.callProcedure('Outer', []);

    const outerIdx = lines.indexOf(2); // Call Inner() in Outer
    const innerIdx = lines.indexOf(5); // x = 99 in Inner
    assert.ok(outerIdx >= 0, 'Outer line found');
    assert.ok(innerIdx >= 0, 'Inner line found');
    assert.strictEqual(depths[outerIdx], 1, 'Outer depth = 1');
    assert.strictEqual(depths[innerIdx], 2, 'Inner depth = 2');
    console.log('[PASS] Call depth correct: Outer=1, Inner=2');
}

// 5. フック内で env から変数が読める
{
    const src = [
        'Sub Main()',
        '  Dim x As Long',  // x 宣言
        '  x = 42',         // x に代入後にフック
        '  x = 99',         // この時点で x=42 がフック環境に見える
        'End Sub',
    ].join('\n');

    const captured: number[] = [];
    const hook: DebugHook = {
        onBeforeStatement(line, _depth, env) {
            if (line === 4) {
                // x = 99 の前（x = 42 実行済み）
                const vars = env.getLocalVariables();
                if (vars.has('x')) captured.push(vars.get('x'));
            }
        },
    };

    const ast = new Parser(new Lexer(src).tokenize()).parse();
    const evaluator = new Evaluator(() => { /* suppress */ });
    evaluator.setDebugHook(hook);
    evaluator.setSourceModule('Module1');
    evaluator.evaluateModule(ast);
    evaluator.resolveIdentifiers([{ ast, moduleName: 'Module1' }]);
    evaluator.callProcedure('Main', []);

    assert.strictEqual(captured.length, 1, 'hook captured x before line 4');
    assert.strictEqual(captured[0], 42, 'x = 42 at line 4 entry');
    console.log('[PASS] Hook env reflects variable value correctly');
}

// 6. setDebugHook(null) でフックを外せる
{
    const src = 'Sub Main()\n  x = 1\nEnd Sub';
    const { evaluator, lines } = makeEvaluator(src);
    evaluator.setDebugHook(null);
    evaluator.callProcedure('Main', []);
    assert.strictEqual(lines.length, 0, 'no hook after setDebugHook(null)');
    console.log('[PASS] setDebugHook(null) disables hook');
}

// 7. コールスタック情報がフックに渡る
{
    const src = [
        'Sub Outer()',
        '  Call Inner()',
        'End Sub',
        'Sub Inner()',
        '  x = 1',
        'End Sub',
    ].join('\n');

    const capturedStacks: ReadonlyArray<{ name: string; moduleName: string; line: number }>[] = [];
    const hook: DebugHook = {
        onBeforeStatement(_line, _depth, _env, callStack) {
            capturedStacks.push(callStack);
        },
    };

    const ast = new Parser(new Lexer(src).tokenize()).parse();
    const evaluator = new Evaluator(() => { /* suppress */ });
    evaluator.setDebugHook(hook);
    evaluator.setSourceModule('Module1');
    evaluator.evaluateModule(ast);
    evaluator.resolveIdentifiers([{ ast, moduleName: 'Module1' }]);
    evaluator.callProcedure('Outer', []);

    const innerStack = capturedStacks.find(s => s.length === 2);
    assert.ok(innerStack !== undefined, 'found 2-frame stack');
    assert.strictEqual(innerStack![0].name, 'Outer', 'bottom frame is Outer');
    assert.strictEqual(innerStack![1].name, 'Inner', 'top frame is Inner');
    console.log('[PASS] Call stack passed to hook correctly');
}

console.log('\n✅ Debug Hook: 全テスト通過');
