/**
 * DefUseAnalyzer (src/engine/def-use-analyzer.ts) のテスト
 * 行範囲を指定して inputs / outputs / locals を算出する
 */
import { Lexer } from '../../src/engine/lexer';
import { Parser } from '../../src/engine/parser';
import { analyzeDefUse } from '../../src/engine/def-use-analyzer';
import { assert } from '../../test-libs/test-runner';

function analyze(code: string, procName: string, startLine: number, endLine: number) {
    const tokens = new Lexer(code).tokenize();
    const ast = new Parser(tokens).parse();
    const proc = (ast.body as any[]).find(
        s => s.type === 'ProcedureDeclaration' && s.name.name === procName
    );
    if (!proc) throw new Error(`Procedure ${procName} not found`);
    return analyzeDefUse(proc, startLine, endLine);
}

// ─── 1. 基本的な input / output / local の分類 ───────────────────────────────
{
    // 1行目: Sub Test()
    // 2行目:     Dim total As Long
    // 3行目:     Dim count As Long
    // 4行目:     total = 0
    // 5行目:     count = 0
    // 6行目:     Dim i As Long
    // 7行目:     For i = 1 To 10
    // 8行目:         total = total + i
    // 9行目:         count = count + 1
    // 10行目:    Next i
    // 11行目:    MsgBox total
    // 12行目: End Sub
    const code = [
        'Sub Test()',
        '    Dim total As Long',
        '    Dim count As Long',
        '    total = 0',
        '    count = 0',
        '    Dim i As Long',
        '    For i = 1 To 10',
        '        total = total + i',
        '        count = count + 1',
        '    Next i',
        '    MsgBox total',
        'End Sub',
    ].join('\n');

    // 範囲: For ループ (行7〜10)
    const r = analyze(code, 'Test', 7, 10);

    assert.strictEqual(r.inputs.includes('total'),  true, 'total は input (範囲前に定義、範囲内で使用)');
    assert.strictEqual(r.inputs.includes('count'),  true, 'count は input (範囲前に定義、範囲内で使用)');
    assert.strictEqual(r.outputs.includes('total'), true, 'total は output (範囲内で定義、範囲後で使用)');
    assert.strictEqual(r.outputs.includes('count'), false, 'count は output でない (範囲後で使用されない)');
    assert.strictEqual(r.locals.includes('i'),      true, 'i は local (範囲内だけで完結)');
    console.log('[PASS] 基本的な input / output / local の分類');
}

// ─── 2. パラメーターは input として扱われる ──────────────────────────────────
{
    const code = [
        'Sub Test(x As Long, y As Long)',
        '    Dim result As Long',
        '    result = x + y',
        '    MsgBox result',
        'End Sub',
    ].join('\n');

    // 範囲: 3行目のみ（result = x + y）
    const r = analyze(code, 'Test', 3, 3);

    assert.strictEqual(r.inputs.includes('x'),      true, 'パラメーター x は input');
    assert.strictEqual(r.inputs.includes('y'),      true, 'パラメーター y は input');
    assert.strictEqual(r.outputs.includes('result'), true, 'result は output (後で使用)');
    console.log('[PASS] パラメーターは input');
}

// ─── 3. 範囲内だけで完結するローカル変数 ─────────────────────────────────────
{
    const code = [
        'Sub Test()',
        '    Dim a As Long',
        '    a = 1',
        '    Dim tmp As Long',
        '    tmp = a * 2',     // line 5: def tmp, use a
        '    a = tmp + 1',     // line 6: def a,   use tmp
        '    MsgBox a',        // line 7: use a
        'End Sub',
    ].join('\n');

    // 範囲: 行5〜6 (tmp の定義と使用が両方収まる)
    const r = analyze(code, 'Test', 5, 6);

    assert.strictEqual(r.inputs.includes('a'),    true,  'a は input (範囲前に定義、範囲内で使用)');
    assert.strictEqual(r.outputs.includes('a'),   true,  'a は output (範囲内で定義、範囲後で使用)');
    assert.strictEqual(r.locals.includes('tmp'),  true,  'tmp は local (範囲内だけで完結)');
    assert.strictEqual(r.outputs.includes('tmp'), false, 'tmp は output でない (範囲後で使用されない)');
    console.log('[PASS] ローカル変数の分類');
}

// ─── 4. ForEach ループ変数は local ───────────────────────────────────────────
{
    const code = [
        'Sub Test()',
        '    Dim col As Variant',
        '    Dim total As Long',
        '    For Each col In Array(1, 2, 3)',
        '        total = total + col',
        '    Next col',
        '    MsgBox total',
        'End Sub',
    ].join('\n');

    // 範囲: 行4〜6 (ForEach ループ)
    const r = analyze(code, 'Test', 4, 6);

    assert.strictEqual(r.inputs.includes('total'),  true, 'total は input');
    assert.strictEqual(r.outputs.includes('total'), true, 'total は output');
    assert.strictEqual(r.locals.includes('col'),    true, 'col (ForEach 変数) は local');
    console.log('[PASS] ForEach ループ変数は local');
}

// ─── 5. 空の範囲（変数なし）───────────────────────────────────────────────────
{
    const code = [
        'Sub Test()',
        '    Dim x As Long',
        '    x = 1',
        'End Sub',
    ].join('\n');

    const r = analyze(code, 'Test', 100, 200);

    assert.strictEqual(r.inputs.length,  0, '空範囲: inputs は空');
    assert.strictEqual(r.outputs.length, 0, '空範囲: outputs は空');
    assert.strictEqual(r.locals.length,  0, '空範囲: locals は空');
    console.log('[PASS] 空の範囲');
}

// ─── 6. If ブロック内の変数 ────────────────────────────────────────────────────
{
    const code = [
        'Sub Test()',
        '    Dim flag As Boolean',
        '    Dim result As Long',
        '    flag = True',
        '    If flag Then',
        '        result = 42',
        '    End If',
        '    MsgBox result',
        'End Sub',
    ].join('\n');

    // 範囲: If ブロック (行5〜7)
    const r = analyze(code, 'Test', 5, 7);

    assert.strictEqual(r.inputs.includes('flag'),    true, 'flag は input');
    assert.strictEqual(r.outputs.includes('result'), true, 'result は output');
    console.log('[PASS] If ブロック内の変数');
}

// ─── 7. 結果はソート済み ─────────────────────────────────────────────────────
{
    const code = [
        'Sub Test()',
        '    Dim c As Long',
        '    Dim a As Long',
        '    Dim b As Long',
        '    c = 1',      // line 5: def c
        '    a = 2',      // line 6: def a
        '    b = 3',      // line 7: def b
        '    MsgBox a + b + c',  // line 8: use a, b, c
        'End Sub',
    ].join('\n');

    const r = analyze(code, 'Test', 5, 7);

    // c, a, b はすべて範囲内で定義・範囲後で使用 → outputs
    assert.strictEqual(r.outputs.length, 3, 'outputs は 3 件');
    assert.strictEqual(r.outputs[0] <= r.outputs[1], true, 'outputs はソート済み (0<=1)');
    assert.strictEqual(r.outputs[1] <= r.outputs[2], true, 'outputs はソート済み (1<=2)');
    assert.deepStrictEqual(r.outputs, ['a', 'b', 'c'], 'outputs = [a, b, c]');
    console.log('[PASS] 結果はソート済み');
}

// ─── 8. CFG 精密化: If 分岐の両パスで定義 → 確実な output ────────────────────
{
    // 線形スキャンでも output になるが、CFG では到達保証を解析できる
    const code = [
        'Sub Test()',
        '    Dim result As Long',
        '    Dim flag As Boolean',
        '    flag = True',
        '    If flag Then',            // line 5
        '        result = 1',          // line 6
        '    Else',
        '        result = 2',          // line 8
        '    End If',                  // line 9
        '    MsgBox result',           // line 10
        'End Sub',
    ].join('\n');

    const r = analyze(code, 'Test', 5, 9);
    assert.strictEqual(r.outputs.includes('result'), true, 'CFG: If/Else 両パスの定義 → output');
    assert.strictEqual(r.inputs.includes('flag'),    true, 'CFG: flag は input');
    console.log('[PASS] CFG: If/Else 両分岐で定義された変数は output');
}

// ─── 9. CFG 精密化: 条件分岐内で read-modify-write → input かつ output ────────
{
    // x = x + 1 はループ内で "前の値を読んで書く" → input AND output
    const code = [
        'Sub Test()',
        '    Dim x As Long',
        '    Dim flag As Boolean',
        '    x = 0',
        '    flag = True',
        '    If flag Then',            // line 6
        '        x = x + 1',          // line 7: x を読んで書く
        '    End If',                  // line 8
        '    MsgBox x',               // line 9
        'End Sub',
    ].join('\n');

    const r = analyze(code, 'Test', 6, 8);
    // x は範囲内で定義され、範囲後で使用される → output
    assert.strictEqual(r.outputs.includes('x'),   true, 'CFG: read-modify-write の x は output');
    // x は範囲前に定義（x=0）され、範囲内で読まれる → input
    assert.strictEqual(r.inputs.includes('x'),    true, 'CFG: read-modify-write の x は input でもある');
    assert.strictEqual(r.inputs.includes('flag'), true, 'CFG: flag は input');
    console.log('[PASS] CFG: read-modify-write は input かつ output');
}

// ─── 10. CFG 精密化: ループ内の変数で範囲外に出ない ─────────────────────────
{
    // ループ内で tmp を計算するが tmp は範囲外で使わない → local
    const code = [
        'Sub Test()',
        '    Dim i As Long',
        '    Dim result As Long',
        '    Dim tmp As Long',
        '    For i = 1 To 5',         // line 5
        '        tmp = i * i',        // line 6
        '        result = result + tmp', // line 7
        '    Next i',                 // line 8
        '    MsgBox result',
        'End Sub',
    ].join('\n');

    const r = analyze(code, 'Test', 5, 8);
    assert.strictEqual(r.locals.includes('i'),      true, 'CFG: ループカウンター i は local');
    assert.strictEqual(r.locals.includes('tmp'),    true, 'CFG: 中間変数 tmp は local');
    assert.strictEqual(r.inputs.includes('result'), true, 'CFG: result は input（前イテレーション依存）');
    assert.strictEqual(r.outputs.includes('result'),true, 'CFG: result は output（後で使用）');
    console.log('[PASS] CFG: ループ内中間変数は local、蓄積変数は input+output');
}

console.log('\n✅ DefUseAnalyzer: 全テスト通過');
