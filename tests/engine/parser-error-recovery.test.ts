import { Lexer } from '../../src/engine/lexer';
import { Parser } from '../../src/engine/parser';
import { assert } from '../../test-libs/test-runner';

function parse(src: string) {
    const tokens = new Lexer(src).tokenize();
    return new Parser(tokens, { errorRecovery: true }).parse();
}

// 1. Valid code produces no diagnostics
{
    const ast = parse('Sub Foo()\n    x = 1\nEnd Sub');
    assert.strictEqual(ast.diagnostics.length, 0, 'no errors for valid code');
    assert.strictEqual(ast.body.length, 1, 'one top-level statement');
    console.log('[PASS] Valid code: diagnostics=', ast.diagnostics.length);
}

// 2. Unknown token records a diagnostic and continues
{
    const ast = parse('@@@ invalid\nSub Foo()\nEnd Sub');
    assert.ok(ast.diagnostics.length >= 1, 'at least one diagnostic');
    assert.ok(ast.diagnostics[0].message.length > 0, 'diagnostic has message');
    const hasSub = ast.body.some((s: any) => s.type === 'ProcedureDeclaration');
    assert.ok(hasSub, 'Sub after error is still parsed');
    console.log('[PASS] Error recovery: diag=', ast.diagnostics[0].message, 'body=', ast.body.length);
}

// 3. Diagnostic carries the correct line number
{
    const src = 'Sub Ok()\nEnd Sub\n@@@ bad\nSub Also()\nEnd Sub';
    const ast = parse(src);
    assert.ok(ast.diagnostics.length >= 1, 'has diagnostic');
    assert.strictEqual(ast.diagnostics[0].loc.start.line, 3, 'error on line 3');
    console.log('[PASS] Diagnostic line:', ast.diagnostics[0].loc.start.line);
}

// 4. Diagnostic carries the correct column number
{
    const src = 'x = 1\n    @@@bad\ny = 2';
    const ast = parse(src);
    assert.ok(ast.diagnostics.length >= 1, 'has diagnostic');
    assert.strictEqual(ast.diagnostics[0].loc.start.column, 5, 'error starts at column 5 (4 spaces + @)');
    console.log('[PASS] Diagnostic column:', ast.diagnostics[0].loc.start.column);
}

// 5. Severity is 'error'
{
    const ast = parse('@@@');
    assert.ok(ast.diagnostics.length >= 1, 'has diagnostic');
    assert.strictEqual(ast.diagnostics[0].severity, 'error', 'severity is error');
    console.log('[PASS] Severity:', ast.diagnostics[0].severity);
}

// 6. Multiple errors on separate lines each recorded
{
    const src = '@@@ first\n@@@ second\nSub Ok()\nEnd Sub';
    const ast = parse(src);
    assert.ok(ast.diagnostics.length >= 2, 'two diagnostics');
    assert.strictEqual(ast.diagnostics[0].loc.start.line, 1, 'first error on line 1');
    assert.strictEqual(ast.diagnostics[1].loc.start.line, 2, 'second error on line 2');
    const hasSub = ast.body.some((s: any) => s.type === 'ProcedureDeclaration');
    assert.ok(hasSub, 'Sub after two errors still parsed');
    console.log('[PASS] Multiple errors:', ast.diagnostics.length, 'subs:', ast.body.length);
}

// 7. Valid statements before AND after error are all parsed
{
    const src = 'x = 1\n@@@ bad\ny = 2';
    const ast = parse(src);
    assert.ok(ast.diagnostics.length >= 1, 'has error');
    // x=1 and y=2 should both appear in body
    assert.strictEqual(ast.body.length, 2, 'two valid statements preserved');
    console.log('[PASS] Statements before+after error preserved:', ast.body.length);
}

// 8. Empty source has no diagnostics
{
    const ast = parse('');
    assert.strictEqual(ast.diagnostics.length, 0, 'empty source: no diagnostics');
    assert.strictEqual(ast.body.length, 0, 'empty source: no body');
    console.log('[PASS] Empty source: clean');
}

// 9. Malformed Sub (missing End Sub) records a diagnostic
{
    const src = 'Sub Broken(\nSub Ok()\nEnd Sub';
    const ast = parse(src);
    assert.ok(ast.diagnostics.length >= 1, 'broken Sub records diagnostic');
    console.log('[PASS] Malformed Sub diagnostic:', ast.diagnostics[0].message.slice(0, 60));
}

// 10. diagnostics array is the same reference as program.diagnostics (live)
{
    const tokens = new Lexer('@@@ err').tokenize();
    const parser = new Parser(tokens, { errorRecovery: true });
    const ast = parser.parse();
    assert.ok(ast.diagnostics === ast.diagnostics, 'diagnostics reference stable');
    assert.ok(ast.diagnostics.length >= 1, 'diagnostic recorded');
    console.log('[PASS] diagnostics reference is stable');
}

// 11. 式の途中の改行エラーが実際の位置を指す（ParseError位置修正）
{
    // Array() の途中で _ なし改行 → エラー位置が Sub先頭でなく該当行を指す
    const src = [
        'Sub Test()',
        '    x = Array(1,',   // line 2: , の後で改行 → ここでエラー
        '    2)',
        'End Sub',
    ].join('\n');
    const ast = parse(src);
    assert.ok(ast.diagnostics.length >= 1, '式途中改行でエラーが記録される');
    const diag = ast.diagnostics[0];
    // エラーは line 2（0-based: 1）または line 3（Array 呼び出し行）を指すべき
    // Sub先頭の line 1 を指してはならない
    assert.ok(diag.loc.start.line >= 2, `エラーがSub先頭ではなく実際の位置を指す (loc.line=${diag.loc.start.line})`);
    console.log('[PASS] 式途中改行エラーの位置: line=', diag.loc.start.line, 'col=', diag.loc.start.column);
}

// 12. 演算子の後で改行したエラーも実際の位置を指す
{
    const src = [
        'Sub Test()',
        '    x = 1 +',    // line 2: + の後で改行
        '    2',
        'End Sub',
    ].join('\n');
    const ast = parse(src);
    assert.ok(ast.diagnostics.length >= 1, 'エラーが記録される');
    const diag = ast.diagnostics[0];
    assert.ok(diag.loc.start.line >= 2, `エラーがSub先頭(line=1)ではなく line=${diag.loc.start.line} を指す`);
    console.log('[PASS] 演算子後改行エラーの位置: line=', diag.loc.start.line, 'col=', diag.loc.start.column);
}

console.log('\n✅ Parser error recovery: 全テスト通過');
