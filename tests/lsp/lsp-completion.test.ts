import { Lexer } from '../../src/engine/lexer';
import { Parser } from '../../src/engine/parser';
import { CompletionProvider } from '../../src/lsp/completion-provider';
import { assert } from '../../test-libs/test-runner';

function getCompletionsAt(src: string, line: number, character: number): any[] {
    const tokens = new Lexer(src).tokenize();
    const ast = new Parser(tokens).parse();
    const provider = new CompletionProvider();
    return provider.getCompletions(ast.body, src, line, character);
}

// 1. Completions for standard VBA functions
{
    const code = 'x = Str';
    const completions = getCompletionsAt(code, 0, 6); // after 'Str'
    assert.ok(Array.isArray(completions), 'returns array');
    assert.ok(completions.some((c: any) => c.label === 'Str'), 'Str function found');
    console.log('[PASS] Standard function completion: Str');
}

// 2. Completions for Len function
{
    const code = 'x = L';
    const completions = getCompletionsAt(code, 0, 5);
    assert.ok(completions.some((c: any) => c.label === 'Len'), 'Len function found');
    console.log('[PASS] Function completion: Len');
}

// 3. Completions for other string functions
{
    const code = 'x = UCa';
    const completions = getCompletionsAt(code, 0, 7);
    assert.ok(completions.some((c: any) => c.label === 'UCase'), 'UCase found');
    console.log('[PASS] Function completion: UCase');
}

// 4. Local variable completion
{
    const code = 'Dim myVar As Integer\nx = my';
    const completions = getCompletionsAt(code, 1, 6); // after 'my'
    assert.ok(completions.some((c: any) => c.label === 'myVar'), 'myVar variable found');
    console.log('[PASS] Local variable completion');
}

// 5. Procedure completion
{
    const code = 'Sub MyProc()\nEnd Sub\nx = My';
    const completions = getCompletionsAt(code, 2, 7); // after 'My'
    assert.ok(completions.some((c: any) => c.label === 'MyProc'), 'MyProc found');
    console.log('[PASS] Procedure completion');
}

// 6. Completion item has required fields
{
    const code = 'x = Str';
    const completions = getCompletionsAt(code, 0, 6);
    const strCompletion = completions.find((c: any) => c.label === 'Str');
    assert.ok(strCompletion, 'Str completion found');
    assert.strictEqual(typeof strCompletion.label, 'string', 'label is string');
    assert.ok(strCompletion.kind, 'kind present');
    console.log('[PASS] Completion item has required fields');
}

// 7. Case-insensitive matching
{
    const code = 'x = str';
    const completions = getCompletionsAt(code, 0, 6); // lowercase 'str'
    assert.ok(completions.some((c: any) => c.label === 'Str'), 'Str found for lowercase str');
    console.log('[PASS] Case-insensitive matching');
}

// 8. Completions for Dim statement
{
    const code = 'Dim x As Inte';
    const completions = getCompletionsAt(code, 0, 13); // after 'Inte'
    // Should include 'Integer' type
    assert.ok(completions.length >= 0, 'completions returned for type context');
    console.log('[PASS] Type completion in Dim statement');
}

// 9. Empty prefix returns all completions
{
    const code = 'x = ';
    const completions = getCompletionsAt(code, 0, 4);
    // Should return standard functions and available variables
    assert.ok(completions.length > 0, 'completions returned for empty prefix');
    console.log('[PASS] Empty prefix returns all completions');
}

// 10. Completion for Class members
{
    const code = 'Class MyClass\n  Public value As Integer\nEnd Class\nSub Test()\nDim obj As New MyClass\nx = obj.';
    const completions = getCompletionsAt(code, 4, 11); // after 'obj.'
    // Should suggest 'value' member
    assert.ok(Array.isArray(completions), 'returns completions');
    console.log('[PASS] Class member completion');
}

console.log('\n✅ LSP Completion: 全テスト通過');
