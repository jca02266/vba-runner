import { Lexer } from '../../src/engine/lexer';
import { Parser } from '../../src/engine/parser';
import { DefinitionProvider } from '../../src/lsp/definition-provider';
import { assert } from '../../test-libs/test-runner';

function getDefinitionAt(src: string, line: number, character: number): any {
    const tokens = new Lexer(src).tokenize();
    const ast = new Parser(tokens).parse();
    const provider = new DefinitionProvider();
    provider.setDocumentUri('file:///test.vba');
    return provider.getDefinition(ast.body, src, line, character);
}

// 1. Definition of Sub procedure
{
    const code = 'Sub Foo()\nEnd Sub';
    const def = getDefinitionAt(code, 0, 4); // 'Foo'
    assert.ok(def, 'definition returned');
    assert.strictEqual(def.uri, 'file:///test.vba', 'uri set');
    assert.ok(def.range, 'range present');
    assert.strictEqual(def.range.start.line, 0, 'starts at line 0');
    console.log('[PASS] Sub definition found');
}

// 2. Definition of Function procedure
{
    const code = 'Function Bar() As Long\nEnd Function';
    const def = getDefinitionAt(code, 0, 9); // 'Bar'
    assert.ok(def, 'definition returned');
    assert.ok(def.range, 'range present');
    console.log('[PASS] Function definition found');
}

// 3. Definition of Class
{
    const code = 'Class MyClass\nEnd Class';
    const def = getDefinitionAt(code, 0, 6); // 'MyClass'
    assert.ok(def, 'definition returned');
    assert.ok(def.range, 'range present');
    console.log('[PASS] Class definition found');
}

// 4. Non-existent identifier returns no definition
{
    const code = 'Sub Foo()\nEnd Sub';
    const def = getDefinitionAt(code, 0, 0); // before 'Sub'
    assert.strictEqual(def, null, 'no definition for non-existent identifier');
    console.log('[PASS] Non-existent identifier: no definition');
}

// 5. Definition with range covering the name
{
    const code = 'Sub TestProc()\nEnd Sub';
    const def = getDefinitionAt(code, 0, 4); // 'TestProc'
    assert.ok(def.range.start, 'range.start present');
    assert.ok(def.range.end, 'range.end present');
    // Range should cover 'TestProc'
    assert.ok(def.range.end.character > def.range.start.character, 'range spans the name');
    console.log('[PASS] Definition range covers identifier');
}

// 6. Definition of variable declaration
{
    const code = 'Dim x As Integer';
    const def = getDefinitionAt(code, 0, 4); // 'x'
    assert.ok(def, 'definition returned');
    assert.ok(def.range, 'range present');
    console.log('[PASS] Variable definition found');
}

// 7. Multi-line: definition on line 1
{
    const code = 'x = 1\nSub Bar()\nEnd Sub';
    const def = getDefinitionAt(code, 1, 4); // 'Bar' on line 1
    assert.ok(def, 'definition returned');
    assert.strictEqual(def.range.start.line, 1, 'definition on line 1');
    console.log('[PASS] Multi-line definition');
}

// 8. Definition of Event
{
    const code = 'Event MyEvent(x As Integer)\nSub Handler()\nEnd Sub';
    const def = getDefinitionAt(code, 0, 6); // 'MyEvent'
    assert.ok(def, 'definition returned');
    assert.ok(def.range, 'range present');
    console.log('[PASS] Event definition found');
}

// 9. Empty source returns no definition
{
    const code = '';
    const def = getDefinitionAt(code, 0, 0);
    assert.strictEqual(def, null, 'empty source: no definition');
    console.log('[PASS] Empty source: no definition');
}

// 10. URI is set correctly
{
    const code = 'Sub Foo()\nEnd Sub';
    const def = getDefinitionAt(code, 0, 4);
    assert.ok(def, 'definition returned');
    assert.strictEqual(def.uri, 'file:///test.vba', 'uri matches');
    console.log('[PASS] URI set correctly');
}

console.log('\n✅ LSP Definition: 全テスト通過');
