import { Lexer } from '../../src/engine/lexer';
import { Parser, ParseDiagnostic } from '../../src/engine/parser';
import { assert } from '../../test-libs/test-runner';

function parse(src: string) {
    const tokens = new Lexer(src).tokenize();
    return new Parser(tokens).parse();
}

function toDiagnostic(diag: ParseDiagnostic): any {
    return {
        range: {
            start: { line: diag.loc.start.line - 1, character: diag.loc.start.column - 1 },
            end: { line: diag.loc.end.line - 1, character: diag.loc.end.column - 1 }
        },
        severity: diag.severity === 'error' ? 1 : 2,
        message: diag.message,
        source: 'vba-runner'
    };
}

// 1. Valid code produces no diagnostics
{
    const ast = parse('Sub Foo()\n    x = 1\nEnd Sub');
    assert.strictEqual(ast.diagnostics.length, 0, 'valid code: no diagnostics');
    console.log('[PASS] Valid code: no diagnostics');
}

// 2. Parse error converts to LSP Diagnostic with correct range
{
    const ast = parse('@@@ bad');
    assert.strictEqual(ast.diagnostics.length, 1, 'one diagnostic');
    const diag = toDiagnostic(ast.diagnostics[0]);
    assert.strictEqual(diag.range.start.line, 0, 'error on line 0 (0-based)');
    assert.strictEqual(diag.range.start.character, 0, 'error at column 0 (0-based)');
    assert.strictEqual(diag.severity, 1, 'severity 1 = error');
    assert.ok(diag.message.length > 0, 'message present');
    assert.strictEqual(diag.source, 'vba-runner', 'source set');
    console.log('[PASS] Diagnostic range and severity correct');
}

// 3. Line/column conversion: 1-based (parser) → 0-based (LSP)
{
    const src = 'x = 1\n    @@@\ny = 2';
    const ast = parse(src);
    assert.strictEqual(ast.diagnostics.length, 1, 'one diagnostic');
    const d = ast.diagnostics[0];
    assert.strictEqual(d.loc.start.line, 2, 'parser: line 2 (1-based)');
    assert.strictEqual(d.loc.start.column, 5, 'parser: column 5 (1-based)');

    const lsp = toDiagnostic(d);
    assert.strictEqual(lsp.range.start.line, 1, 'LSP: line 1 (0-based)');
    assert.strictEqual(lsp.range.start.character, 4, 'LSP: character 4 (0-based)');
    console.log('[PASS] Line/column conversion 1-based → 0-based');
}

// 4. Multiple diagnostics all converted
{
    const ast = parse('@@@ first\n@@@ second');
    assert.strictEqual(ast.diagnostics.length, 2, 'two diagnostics');
    const lsp1 = toDiagnostic(ast.diagnostics[0]);
    const lsp2 = toDiagnostic(ast.diagnostics[1]);
    assert.strictEqual(lsp1.range.start.line, 0, 'first error on line 0');
    assert.strictEqual(lsp2.range.start.line, 1, 'second error on line 1');
    console.log('[PASS] Multiple diagnostics converted');
}

// 5. Error message is preserved
{
    const ast = parse('@@@ token');
    assert.strictEqual(ast.diagnostics.length, 1, 'one diagnostic');
    const lsp = toDiagnostic(ast.diagnostics[0]);
    assert.ok(lsp.message.includes('Unknown token'), 'error message preserved');
    console.log('[PASS] Error message preserved:', lsp.message.slice(0, 40));
}

console.log('\n✅ LSP Diagnostics: 全テスト通過');
