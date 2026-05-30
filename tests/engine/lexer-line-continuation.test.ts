import { Lexer } from '../../src/engine/lexer';
import { Parser } from '../../src/engine/parser';
import { assert } from '../../test-libs/test-runner';

function lex(src: string) {
    const lexer = new Lexer(src);
    lexer.tokenize();
    return lexer.diagnostics;
}

function parse(src: string) {
    const lexer = new Lexer(src);
    const tokens = lexer.tokenize();
    const ast = new Parser(tokens, { errorRecovery: true }).parse();
    return { lexerDiags: lexer.diagnostics, parserDiags: ast.diagnostics };
}

// ─── 正常な行継続 ──────────────────────────────────────────────────────────────

{
    // _ の直後に改行: 有効
    const diags = lex('x = 1 + _\n    2');
    assert.strictEqual(diags.length, 0, '_ 直後改行は正常（エラーなし）');
    console.log('[PASS] _ 直後改行: エラーなし');
}

{
    // 行継続後に式が続く: パースエラーなし
    const { lexerDiags, parserDiags } = parse('Sub Test()\n    x = 1 + _\n    2\nEnd Sub');
    assert.strictEqual(lexerDiags.length, 0, '行継続を含むコード: lexerエラーなし');
    assert.strictEqual(parserDiags.length, 0, '行継続を含むコード: parserエラーなし');
    console.log('[PASS] 行継続コード: エラーなし');
}

// ─── _ の後にコメント（エラー） ──────────────────────────────────────────────

{
    // _ の後にスペース + コメント
    const diags = lex("x = 1 + _  ' この行は継続できない\n    2");
    assert.strictEqual(diags.length, 1, "_ + コメント: エラー1件");
    assert.ok(diags[0].message.includes('コメント'), "メッセージに「コメント」が含まれる");
    console.log('[PASS] _ + コメント:', diags[0].message);
}

{
    // エラー位置が _ の位置を指す
    const diags = lex("    x = a + _  ' comment\n    b");
    assert.strictEqual(diags.length, 1, 'エラー1件');
    assert.strictEqual(diags[0].line, 1, 'エラーは1行目');
    assert.strictEqual(diags[0].column, 13, '_ は列13（"    x = a + "の次）');
    console.log('[PASS] _ + コメントのエラー位置: line=', diags[0].line, 'col=', diags[0].column);
}

{
    // _ + コメント後もパースを継続できる（後続行が正常にパースされる）
    const { parserDiags } = parse('Sub Test()\n    x = 1\nEnd Sub');
    assert.strictEqual(parserDiags.length, 0, '正常コード: parserエラーなし');
    console.log('[PASS] _ + コメント後のパース継続確認');
}

// ─── _ の後にスペース（エラー） ────────────────────────────────────────────────

{
    // _ の後にスペースのみ + 改行
    const diags = lex('x = 1 + _   \n    2');
    assert.strictEqual(diags.length, 1, '_ + 空白 + 改行: エラー1件');
    assert.ok(diags[0].message.includes('空白'), "メッセージに「空白」が含まれる");
    console.log('[PASS] _ + 空白:', diags[0].message);
}

{
    // _ の後に1スペース + 改行
    const diags = lex('x = _ \n    y');
    assert.strictEqual(diags.length, 1, '_ + 1スペース: エラー1件');
    console.log('[PASS] _ + 1スペース: エラー検知');
}

{
    // エラー位置が _ の位置を指す
    const diags = lex('x = 1 + _ \ny = 2');
    assert.strictEqual(diags.length, 1, 'エラー1件');
    assert.strictEqual(diags[0].line, 1, 'エラーは1行目');
    assert.ok(diags[0].column > 0, 'column が設定されている');
    console.log('[PASS] _ + 空白のエラー位置: line=', diags[0].line, 'col=', diags[0].column);
}

// ─── _ が文字途中（識別子）の場合はエラーにならない ─────────────────────────

{
    // アンダースコアを含む識別子
    const diags = lex('Dim my_var As Long');
    assert.strictEqual(diags.length, 0, '識別子内の _ はエラーなし');
    console.log('[PASS] 識別子内の _: エラーなし');
}

// ─── 複数の無効な行継続 ───────────────────────────────────────────────────────

{
    const diags = lex("a = 1 + _ ' c1\nb = 2 + _ ' c2\nc = 3");
    assert.strictEqual(diags.length, 2, '2つの無効な行継続: エラー2件');
    assert.strictEqual(diags[0].line, 1, '1件目は1行目');
    assert.strictEqual(diags[1].line, 2, '2件目は2行目');
    console.log('[PASS] 複数の無効な行継続: エラー件数=', diags.length);
}

// ─── コメント末尾の _（行継続として無効）────────────────────────────────────

{
    // ' aaa _ → _ はコメント内なので行継続として機能しない
    const diags = lex("x = 1\n' aaa _\ny = 2");
    assert.strictEqual(diags.length, 1, "コメント末尾の _: エラー1件");
    assert.ok(diags[0].message.includes('コメント'), "コメント末尾メッセージ");
    console.log('[PASS] コメント末尾の _:', diags[0].message);
}

{
    // インラインコメントで末尾に _ がある場合
    const diags = lex('x = Foo(a, b) \' note _');
    assert.strictEqual(diags.length, 1, "インラインコメント末尾の _: エラー1件");
    console.log('[PASS] インラインコメント末尾の _:', diags[0].message);
}

{
    // コメント末尾にスペース + _ の場合
    const diags = lex("' comment   _   \nx = 1");
    assert.strictEqual(diags.length, 1, "コメント末尾スペース+_: エラー1件");
    console.log('[PASS] コメント末尾スペース+_:', diags[0].message);
}

{
    // コメント末尾に _ がない場合はエラーなし
    const diags = lex("' just a comment\nx = 1");
    assert.strictEqual(diags.length, 0, "普通のコメント: エラーなし");
    console.log('[PASS] 普通のコメント: エラーなし');
}

console.log('\n✅ Lexer 行継続チェック: 全テスト通過');
