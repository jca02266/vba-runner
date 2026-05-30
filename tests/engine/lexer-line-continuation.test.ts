import { Lexer, LexError } from '../../src/engine/lexer';
import { Parser } from '../../src/engine/parser';
import { assert } from '../../test-libs/test-runner';

function lex(src: string) {
    return new Lexer(src).tokenize();
}

function parse(src: string) {
    const lexer = new Lexer(src);
    const tokens = lexer.tokenize();
    const ast = new Parser(tokens, { errorRecovery: true }).parse();
    return ast.diagnostics;
}

function expectLexError(src: string, msgFragment: string, label: string) {
    try {
        lex(src);
        assert.ok(false, `${label}: LexError が投げられなかった`);
    } catch (e: any) {
        assert.ok(e instanceof LexError, `${label}: LexError のインスタンス`);
        assert.ok(e.message.includes(msgFragment), `${label}: メッセージに「${msgFragment}」が含まれる（got: ${e.message}）`);
    }
}

// ─── 正常な行継続 ──────────────────────────────────────────────────────────────

{
    // _ の直後に改行: 有効
    const tokens = lex('x = 1 + _\n    2');
    assert.ok(tokens.length > 0, '_ 直後改行は正常（トークン生成）');
    console.log('[PASS] _ 直後改行: エラーなし');
}

{
    // 行継続後に式が続く: パースエラーなし
    const parserDiags = parse('Sub Test()\n    x = 1 + _\n    2\nEnd Sub');
    assert.strictEqual(parserDiags.length, 0, '行継続を含むコード: parserエラーなし');
    console.log('[PASS] 行継続コード: エラーなし');
}

// ─── _ の後にコメント（LexError） ────────────────────────────────────────────

{
    expectLexError(
        "x = 1 + _  ' この行は継続できない\n    2",
        'コメント',
        '_ + コメント'
    );
    console.log('[PASS] _ + コメント: LexError をスロー');
}

{
    // エラー位置が _ の位置を指す
    try {
        lex("    x = a + _  ' comment\n    b");
        assert.ok(false, 'LexError が投げられなかった');
    } catch (e: any) {
        assert.ok(e instanceof LexError, 'LexError のインスタンス');
        assert.strictEqual(e.line, 1, 'エラーは1行目');
        assert.strictEqual(e.column, 13, '_ は列13（"    x = a + "の次）');
        console.log('[PASS] _ + コメントのエラー位置: line=', e.line, 'col=', e.column);
    }
}

// ─── _ の後にスペース（LexError） ─────────────────────────────────────────────

{
    expectLexError(
        'x = 1 + _   \n    2',
        '空白',
        '_ + 空白 + 改行'
    );
    console.log('[PASS] _ + 空白: LexError をスロー');
}

{
    expectLexError('x = _ \n    y', '空白', '_ + 1スペース');
    console.log('[PASS] _ + 1スペース: LexError をスロー');
}

{
    // エラー位置が _ の位置を指す
    try {
        lex('x = 1 + _ \ny = 2');
        assert.ok(false, 'LexError が投げられなかった');
    } catch (e: any) {
        assert.ok(e instanceof LexError, 'LexError のインスタンス');
        assert.strictEqual(e.line, 1, 'エラーは1行目');
        assert.ok(e.column > 0, 'column が設定されている');
        console.log('[PASS] _ + 空白のエラー位置: line=', e.line, 'col=', e.column);
    }
}

// ─── _ が文字途中（識別子）の場合はエラーにならない ─────────────────────────

{
    const tokens = lex('Dim my_var As Long');
    assert.ok(tokens.length > 0, '識別子内の _ はエラーなし');
    console.log('[PASS] 識別子内の _: エラーなし');
}

// ─── コメント末尾の _（行継続として無効）────────────────────────────────────

{
    // ' aaa _ → _ はコメント内なので行継続として機能しない
    const lexer = new Lexer("x = 1\n' aaa _\ny = 2");
    lexer.tokenize();
    assert.strictEqual(lexer.diagnostics.length, 1, "コメント末尾の _: diagnostics に1件");
    assert.ok(lexer.diagnostics[0].message.includes('コメント'), "コメント末尾メッセージ");
    console.log('[PASS] コメント末尾の _:', lexer.diagnostics[0].message);
}

{
    const lexer = new Lexer('x = Foo(a, b) \' note _');
    lexer.tokenize();
    assert.strictEqual(lexer.diagnostics.length, 1, "インラインコメント末尾の _: diagnostics に1件");
    console.log('[PASS] インラインコメント末尾の _:', lexer.diagnostics[0].message);
}

{
    const lexer = new Lexer("' comment   _   \nx = 1");
    lexer.tokenize();
    assert.strictEqual(lexer.diagnostics.length, 1, "コメント末尾スペース+_: diagnostics に1件");
    console.log('[PASS] コメント末尾スペース+_:', lexer.diagnostics[0].message);
}

{
    const lexer = new Lexer("' just a comment\nx = 1");
    lexer.tokenize();
    assert.strictEqual(lexer.diagnostics.length, 0, "普通のコメント: エラーなし");
    console.log('[PASS] 普通のコメント: エラーなし');
}

console.log('\n✅ Lexer 行継続チェック: 全テスト通過');
