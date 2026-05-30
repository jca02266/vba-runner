import * as assert from 'assert';
import { Lexer } from '../../src/engine/lexer';
import { Parser } from '../../src/engine/parser';
import { inferVariantTypes, buildProcMap, findProcAtLine } from '../../src/lsp/variant-type-inferencer';

function parse(src: string) {
    const tokens = new Lexer(src).tokenize();
    return new Parser(tokens, { errorRecovery: true }).parse().body;
}

function infer(src: string, procName?: string) {
    const stmts = parse(src);
    const procMap = buildProcMap(stmts);
    const proc = procName
        ? procMap.get(procName.toLowerCase())!
        : [...procMap.values()][0];
    const memo = new Map();
    return inferVariantTypes(proc, procMap, memo);
}

// ─── リテラル代入 ─────────────────────────────────────────────────────────────

{
    const hints = infer([
        'Sub Test()',
        '    Dim x',
        '    x = 42',
        'End Sub',
    ].join('\n'));
    const h = hints.find(h => h.varName === 'x');
    assert.strictEqual(h?.inferredType, 'Long', 'x = 42 → Long');
    console.log('[PASS] リテラル整数 → Long');
}

{
    const hints = infer([
        'Sub Test()',
        '    Dim x',
        '    x = 3.14',
        'End Sub',
    ].join('\n'));
    const h = hints.find(h => h.varName === 'x');
    assert.strictEqual(h?.inferredType, 'Double', 'x = 3.14 → Double');
    console.log('[PASS] リテラル小数 → Double');
}

{
    const hints = infer([
        'Sub Test()',
        '    Dim s',
        '    s = "hello"',
        'End Sub',
    ].join('\n'));
    const h = hints.find(h => h.varName === 's');
    assert.strictEqual(h?.inferredType, 'String', 's = "hello" → String');
    console.log('[PASS] 文字列リテラル → String');
}

{
    const hints = infer([
        'Sub Test()',
        '    Dim b',
        '    b = True',
        'End Sub',
    ].join('\n'));
    const h = hints.find(h => h.varName === 'b');
    assert.strictEqual(h?.inferredType, 'Boolean', 'b = True → Boolean');
    console.log('[PASS] Boolean リテラル → Boolean');
}

// ─── 演算式 ───────────────────────────────────────────────────────────────────

{
    const hints = infer([
        'Sub Test()',
        '    Dim x',
        '    x = 1 + 2',
        'End Sub',
    ].join('\n'));
    const h = hints.find(h => h.varName === 'x');
    assert.strictEqual(h?.inferredType, 'Long', '1 + 2 → Long');
    console.log('[PASS] Long + Long → Long');
}

{
    const hints = infer([
        'Sub Test()',
        '    Dim x',
        '    x = "a" & "b"',
        'End Sub',
    ].join('\n'));
    const h = hints.find(h => h.varName === 'x');
    assert.strictEqual(h?.inferredType, 'String', '"a" & "b" → String');
    console.log('[PASS] & 演算 → String');
}

// ─── 明示的型の関数呼び出し（1段） ───────────────────────────────────────────

{
    const hints = infer([
        'Function GetName() As String',
        '    GetName = "test"',
        'End Function',
        'Sub Test()',
        '    Dim x',
        '    x = GetName()',
        'End Sub',
    ].join('\n'), 'Test');
    const h = hints.find(h => h.varName === 'x');
    assert.strictEqual(h?.inferredType, 'String', 'x = GetName() As String → String');
    console.log('[PASS] 関数戻り型（明示: String）→ String');
}

{
    const hints = infer([
        'Function GetCount() As Long',
        '    GetCount = 10',
        'End Function',
        'Sub Test()',
        '    Dim n',
        '    n = GetCount()',
        'End Sub',
    ].join('\n'), 'Test');
    const h = hints.find(h => h.varName === 'n');
    assert.strictEqual(h?.inferredType, 'Long', 'n = GetCount() As Long → Long');
    console.log('[PASS] 関数戻り型（明示: Long）→ Long');
}

// ─── Variant 関数の再帰推論 ───────────────────────────────────────────────────

{
    const hints = infer([
        'Function Compute()',    // 戻り型なし → Variant → ボディを推論
        '    Compute = 99',
        'End Function',
        'Sub Test()',
        '    Dim result',
        '    result = Compute()',
        'End Sub',
    ].join('\n'), 'Test');
    const h = hints.find(h => h.varName === 'result');
    assert.strictEqual(h?.inferredType, 'Long', 'Variant 関数ボディから Long を推論');
    console.log('[PASS] Variant 関数の再帰推論 → Long');
}

// ─── メモ化（同じ関数を複数変数で参照） ─────────────────────────────────────

{
    const stmts = parse([
        'Function GetVal() As Long',
        '    GetVal = 1',
        'End Function',
        'Sub Test()',
        '    Dim a',
        '    Dim b',
        '    a = GetVal()',
        '    b = GetVal()',
        'End Sub',
    ].join('\n'));
    const procMap = buildProcMap(stmts);
    const proc = procMap.get('test')!;
    const memo = new Map<string, any>();
    const hints = inferVariantTypes(proc, procMap, memo);

    assert.ok(memo.has('getval'), 'GetVal がメモ化されている');
    assert.strictEqual(memo.get('getval'), 'Long', 'メモ値が Long');
    assert.strictEqual(hints.filter(h => h.inferredType === 'Long').length, 2, 'a・b 両方 Long');
    console.log('[PASS] メモ化: 同じ関数の2回目はメモから取得');
}

// ─── 推論不能ケース（ヒントなし） ────────────────────────────────────────────

{
    // 未代入の Variant はヒントを出さない
    const hints = infer([
        'Sub Test()',
        '    Dim x',
        '    Debug.Print x',
        'End Sub',
    ].join('\n'));
    const h = hints.find(h => h.varName === 'x');
    assert.strictEqual(h, undefined, '未代入の変数はヒントなし');
    console.log('[PASS] 未代入の Variant → ヒントなし');
}

// ─── findProcAtLine ───────────────────────────────────────────────────────────

{
    const stmts = parse([
        'Sub First()',   // line 1
        '    x = 1',    // line 2
        'End Sub',      // line 3
        'Sub Second()', // line 4
        '    y = 2',    // line 5
        'End Sub',      // line 6
    ].join('\n'));

    const proc1 = findProcAtLine(stmts, 1); // 0-based line 1 → Second?
    const proc2 = findProcAtLine(stmts, 4); // 0-based line 4
    assert.strictEqual(proc1?.name.name.toLowerCase(), 'first',  'line 1 → First');
    assert.strictEqual(proc2?.name.name.toLowerCase(), 'second', 'line 4 → Second');
    console.log('[PASS] findProcAtLine: カーソル行から手続きを特定');
}

console.log('\n✅ Variant 型推論: 全テスト通過');
