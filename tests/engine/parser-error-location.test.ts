/**
 * パースエラーの位置精度テスト
 *
 * 各エラー種別ごとに「エラーが正しい行・列を指すか」を検証する。
 * loc はすべて 1-based。
 *
 * 既知の制限:
 *   E09/E10/E11: 対応するクローズキーワード（End If / Next / Loop）が欠落した
 *     場合、エラーは "End Sub" が現れた行で報告される（開始文の行ではない）。
 *     これはパーサーが混乱した位置を報告する現在の仕様。
 */

import { Lexer, LexError } from '../../src/engine/lexer';
import { Parser } from '../../src/engine/parser';
import { assert } from '../../test-libs/test-runner';

interface Diag { line: number; col: number; msg: string; src: 'lexer' | 'parser' }

function diagnose(src: string): Diag[] {
    const lexer = new Lexer(src);
    let tokens;
    try {
        tokens = lexer.tokenize();
    } catch (e) {
        if (e instanceof LexError) {
            return [{ line: e.line, col: e.column, msg: e.message.replace(/ \(line \d+\)$/, ''), src: 'lexer' }];
        }
        throw e;
    }
    const ast = new Parser(tokens, { errorRecovery: true }).parse();
    return [
        ...lexer.diagnostics.map(d => ({ line: d.line, col: d.column, msg: d.message, src: 'lexer' as const })),
        ...ast.diagnostics.map(d => ({ line: d.loc.start.line, col: d.loc.start.column, msg: d.message, src: 'parser' as const })),
    ];
}

function assertError(label: string, diags: Diag[], expectedLine: number, expectedCol?: number) {
    assert.ok(diags.length >= 1, `${label}: エラーが1件以上あること`);
    const d = diags[0];
    assert.strictEqual(d.line, expectedLine, `${label}: エラーが line=${expectedLine} を指すこと (got ${d.line})`);
    if (expectedCol !== undefined) {
        assert.strictEqual(d.col, expectedCol, `${label}: エラーが col=${expectedCol} を指すこと (got ${d.col})`);
    }
}

// ─── E01: 行頭の不明文字 ─────────────────────────────────────────────────────
{
    const diags = diagnose('Sub Test()\n    @ = 1\nEnd Sub');
    assertError('E01', diags, 2, 5);
    assert.ok(diags[0].msg.includes('@'), 'E01: メッセージに @ を含む');
    console.log('[PASS] E01: 行頭の不明文字 @ → line=2 col=5');
}

// ─── E02: 式の途中の不明文字 ─────────────────────────────────────────────────
{
    const diags = diagnose('Sub Test()\n    x = 1 @ 2\nEnd Sub');
    assertError('E02', diags, 2, 11);
    console.log('[PASS] E02: 式中の不明文字 @ → line=2 col=11');
}

// ─── E03: 演算子 + の後で改行（_ なし）──────────────────────────────────────
{
    const diags = diagnose('Sub Test()\n    x = 1 +\n    2\nEnd Sub');
    assertError('E03', diags, 2);
    assert.ok(diags[0].msg.includes("'_'") || diags[0].msg.includes('Unexpected token'), 'E03: _ または Unexpected token メッセージ');
    console.log('[PASS] E03: + 後の改行エラー → line=2');
}

// ─── E04: カンマの後で改行（_ なし）─────────────────────────────────────────
{
    const diags = diagnose('Sub Test()\n    x = Foo(a,\n    b)\nEnd Sub');
    assertError('E04', diags, 2);
    console.log('[PASS] E04: , 後の改行エラー → line=2');
}

// ─── E05: 開き括弧の後で改行（_ なし）───────────────────────────────────────
{
    const diags = diagnose('Sub Test()\n    x = Foo(\n    a)\nEnd Sub');
    assertError('E05', diags, 2);
    console.log('[PASS] E05: ( 後の改行エラー → line=2');
}

// ─── E06: 代入 = の後で改行（_ なし）────────────────────────────────────────
{
    const diags = diagnose('Sub Test()\n    x =\n    1\nEnd Sub');
    assertError('E06', diags, 2);
    assert.ok(diags[0].msg.includes("'_'") || diags[0].msg.includes('Unexpected token'), 'E06: _ または Unexpected token メッセージ');
    console.log('[PASS] E06: = 後の改行エラー → line=2');
}

// ─── E07: 閉じ括弧 ) が不足 ─────────────────────────────────────────────────
{
    const diags = diagnose('Sub Test()\n    x = Foo(a, b\nEnd Sub');
    assertError('E07', diags, 2);
    assert.ok(diags[0].msg.includes(')'), "E07: ) 不足メッセージ");
    console.log('[PASS] E07: ) 不足 → line=2');
}

// ─── E08: End Sub が欠落 ─────────────────────────────────────────────────────
{
    const diags = diagnose('Sub Test()\n    x = 1\n');
    assertError('E08', diags, 2);
    assert.ok(diags[0].msg.includes('End Sub'), `E08: 'End Sub' メッセージ (got: ${diags[0].msg})`);
    console.log('[PASS] E08: End Sub 欠落 → line=2 でエラー');
}

// ─── E09: End If が欠落 ──────────────────────────────────────────────────────
{
    // If（line 2）が開いたまま End Sub（line 4）に到達する
    // → "Expected 'If' after 'End'" が End Sub の行で報告される
    const src = [
        'Sub Test()',       // line 1
        '    If x > 0 Then',// line 2
        '        x = 1',   // line 3
        'End Sub',         // line 4
    ].join('\n');
    const diags = diagnose(src);
    assertError('E09', diags, 4);
    assert.ok(diags[0].msg.includes('If'), "E09: 'If' 欠落メッセージ");
    console.log('[PASS] E09: End If 欠落 → End Sub 行 (line=4) でエラー');
}

// ─── E10: Next が欠落（For ループ）──────────────────────────────────────────
{
    // For（line 2）の Next が欠落; End Sub（line 4）で検知
    const src = [
        'Sub Test()',         // line 1
        '    For i = 1 To 10',// line 2
        '        x = i',     // line 3
        'End Sub',           // line 4
    ].join('\n');
    const diags = diagnose(src);
    assertError('E10', diags, 4);
    assert.ok(diags[0].msg.includes('Next'), "E10: 'Next' 欠落メッセージ");
    console.log('[PASS] E10: Next 欠落 → End Sub 行 (line=4) でエラー（ハングなし）');
}

// ─── E11: Loop が欠落（Do While ループ）─────────────────────────────────────
{
    // Do While（line 2）の Loop が欠落; End Sub（line 4）で検知
    const src = [
        'Sub Test()',          // line 1
        '    Do While x > 0', // line 2
        '        x = x - 1', // line 3
        'End Sub',            // line 4
    ].join('\n');
    const diags = diagnose(src);
    assertError('E11', diags, 4);
    assert.ok(diags[0].msg.includes('Loop'), "E11: 'Loop' 欠落メッセージ");
    console.log('[PASS] E11: Loop 欠落 → End Sub 行 (line=4) でエラー（ハングなし）');
}

// ─── E12: _ + 空白 + 改行（lexer エラー）────────────────────────────────────
{
    const diags = diagnose('Sub Test()\n    x = 1 + _ \n    2\nEnd Sub');
    assertError('E12', diags, 2, 13);
    assert.strictEqual(diags[0].src, 'lexer', 'E12: lexer が検知');
    assert.ok(diags[0].msg.includes('空白'), 'E12: 空白メッセージ');
    console.log('[PASS] E12: _ + 空白 → line=2 col=13 (lexer)');
}

// ─── E13: _ + コメント + 改行（lexer エラー）────────────────────────────────
{
    const diags = diagnose("Sub Test()\n    x = 1 + _  ' comment\n    2\nEnd Sub");
    assertError('E13', diags, 2, 13);
    assert.strictEqual(diags[0].src, 'lexer', 'E13: lexer が検知');
    assert.ok(diags[0].msg.includes('コメント'), 'E13: コメントメッセージ');
    console.log('[PASS] E13: _ + コメント → line=2 col=13 (lexer)');
}

// ─── E14: 複数行にまたがる Array() 呼び出しで途中1行だけ _ 行継続が欠落 ──────
{
    const src = [
        'Sub Test()',           // line 1
        '    rows = Array( _',  // line 2
        '        Array(1, 2), _',// line 3
        '        Array(3, 4),', // line 4: _ が欠落
        '        Array(5, 6) _',// line 5
        '    )',                // line 6
        'End Sub',             // line 7
    ].join('\n');
    const diags = diagnose(src);
    assertError('E14', diags, 4);
    console.log('[PASS] E14: Array() 途中改行 → line=4 でエラー');
}

// ─── E15: 論理キーワード + コメント + 改行（_ なし）────────────────────────
// `And ' comment \n` はコメントが lexer で除去されるため
// `And \n` と同等になり、And の行でエラーが報告される
{
    const diags = diagnose("Sub Test()\n    x = 1 And ' 右辺がない\n    2\nEnd Sub");
    assertError('E15', diags, 2);
    assert.ok(diags[0].msg.includes("'_'") || diags[0].msg.includes('Unexpected token'), 'E15: _ または Unexpected token メッセージ');
    console.log('[PASS] E15: And + コメント + 改行 → line=2 でエラー');
}

// ─── E16: Or + コメント + 改行（_ なし）────────────────────────────────────
{
    const diags = diagnose("Sub Test()\n    If a Or ' check b\n        b Then x = 1\n    End If\nEnd Sub");
    assertError('E16', diags, 2);
    console.log('[PASS] E16: Or + コメント + 改行 → line=2 でエラー');
}

// ─── E17: End Sub がエラー回復後に連鎖エラーを起こさない ────────────────────
// Sub 内でパースエラーが起きたとき、その Sub の End Sub が
// "Expected procedure name" などの連鎖エラーを出してはならない
{
    const src = [
        'Sub CreateData()',     // line 1
        '    x = Array(1,',    // line 2: ) が欠落 → エラー
        '    2)',               // line 3
        'End Sub',             // line 4: 連鎖エラーが出ないこと
        'Sub Other()',         // line 5
        'End Sub',             // line 6
    ].join('\n');
    const diags = diagnose(src);
    // 最初のエラーは line 2 の ) 欠落
    assert.ok(diags.length >= 1, 'E17: 少なくとも1件のエラー');
    assert.strictEqual(diags[0].line, 2, 'E17: 最初のエラーが line=2');
    // line 4 (End Sub) や line 5 (Sub Other) への連鎖エラーがないこと
    const cascading = diags.slice(1).filter(d => d.line >= 4);
    assert.strictEqual(cascading.length, 0, `E17: End Sub 以降に連鎖エラーなし (got: ${JSON.stringify(cascading)})`);
    console.log('[PASS] E17: End Sub が連鎖エラーを起こさない');
}

console.log('\n✅ パースエラー位置精度テスト: 全テスト通過');
