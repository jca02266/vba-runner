import { Lexer } from '../../src/engine/lexer';
import { Parser } from '../../src/engine/parser';
import { SymbolProvider, SymbolKind } from '../../src/lsp/symbol-provider';
import { assert } from '../../test-libs/test-runner';

function parseAndExtract(src: string): any[] {
    const tokens = new Lexer(src).tokenize();
    const ast = new Parser(tokens).parse();
    const provider = new SymbolProvider();
    provider.setDocumentUri('file:///test.vba');
    return provider.extractSymbols(ast.body);
}

// 1. Sub procedure symbol extracted
{
    const syms = parseAndExtract('Sub Foo()\nEnd Sub');
    assert.strictEqual(syms.length, 1, 'one symbol');
    assert.strictEqual(syms[0].name, 'Foo', 'name is Foo');
    assert.strictEqual(syms[0].kind, SymbolKind.Method, 'kind is Method for Sub');
    console.log('[PASS] Sub procedure extracted');
}

// 2. Function symbol with Function kind
{
    const syms = parseAndExtract('Function Bar() As Long\nEnd Function');
    assert.strictEqual(syms.length, 1, 'one symbol');
    assert.strictEqual(syms[0].name, 'Bar', 'name is Bar');
    assert.strictEqual(syms[0].kind, SymbolKind.Function, 'kind is Function');
    console.log('[PASS] Function extracted with Function kind');
}

// 3. Property symbol with Property kind
{
    const syms = parseAndExtract('Property Get MyProp() As String\nEnd Property');
    assert.strictEqual(syms.length, 1, 'one symbol');
    assert.strictEqual(syms[0].kind, SymbolKind.Property, 'kind is Property');
    console.log('[PASS] Property extracted with Property kind');
}

// 4. Class with members in children array
{
    const code = `
    Class MyClass
        Public Sub Init()
        End Sub
        Public Function GetValue() As Integer
        End Function
    End Class
    `;
    const syms = parseAndExtract(code);
    assert.strictEqual(syms.length, 1, 'one top-level symbol');
    const cls = syms[0];
    assert.strictEqual(cls.name, 'MyClass', 'class name');
    assert.strictEqual(cls.kind, SymbolKind.Class, 'kind is Class');
    assert.ok(cls.children, 'children array exists');
    assert.strictEqual(cls.children!.length, 2, 'class has 2 members');
    assert.strictEqual(cls.children![0].name, 'Init', 'first member is Init');
    assert.strictEqual(cls.children![1].name, 'GetValue', 'second member is GetValue');
    console.log('[PASS] Class members nested in children');
}

// 5. Member symbols have containerName set
{
    const code = `
    Class Foo
        Public Sub Bar()
        End Sub
    End Class
    `;
    const syms = parseAndExtract(code);
    const cls = syms[0];
    assert.ok(cls.children, 'children exist');
    assert.strictEqual(cls.children![0].containerName, 'Foo', 'method has containerName');
    console.log('[PASS] Member containerName set to class name');
}

// 6. Variable declarations extracted as Variable kind
{
    const syms = parseAndExtract('Dim x As Integer\nDim y As String');
    assert.strictEqual(syms.length, 2, 'two symbols');
    assert.strictEqual(syms[0].name, 'x', 'first is x');
    assert.strictEqual(syms[0].kind, SymbolKind.Variable, 'kind is Variable');
    assert.strictEqual(syms[1].name, 'y', 'second is y');
    console.log('[PASS] Variables extracted with Variable kind');
}

// 7. Location has uri and range
{
    const syms = parseAndExtract('Sub Test()\nEnd Sub');
    const sym = syms[0];
    assert.ok(sym.location, 'location present');
    assert.strictEqual(sym.location.uri, 'file:///test.vba', 'uri is set');
    assert.ok(sym.location.range, 'range present');
    assert.ok(typeof sym.location.range.start.line === 'number', 'start.line is number');
    assert.ok(typeof sym.location.range.start.character === 'number', 'start.character is number');
    console.log('[PASS] Location has uri and range');
}

// 8. Line/column converted to 0-based
{
    const code = 'x = 1\nSub Foo()\nEnd Sub';
    const syms = parseAndExtract(code);
    // Sub Foo starts at line 2, column 1 in parser (1-based)
    // Should be line 1, character 0 in LSP (0-based)
    const subSym = syms.find((s: any) => s.name === 'Foo');
    assert.ok(subSym, 'Foo symbol found');
    assert.strictEqual(subSym.location.range.start.line, 1, 'start.line is 1 (0-based)');
    assert.strictEqual(subSym.location.range.start.character, 0, 'start.character is 0 (0-based)');
    console.log('[PASS] Line/column 1-based → 0-based conversion');
}

// 9. Empty source produces no symbols
{
    const syms = parseAndExtract('');
    assert.strictEqual(syms.length, 0, 'empty: no symbols');
    console.log('[PASS] Empty source: no symbols');
}

// 10. Mixed declarations
{
    const code = `
    Dim global As Integer
    Sub Init()
    End Sub
    Class Data
        Public value As Long
    End Class
    Function Process() As Boolean
    End Function
    `;
    const syms = parseAndExtract(code);
    assert.ok(syms.length >= 4, 'multiple symbols extracted');
    assert.ok(syms.some((s: any) => s.name === 'global'), 'global variable');
    assert.ok(syms.some((s: any) => s.name === 'Init'), 'Sub Init');
    assert.ok(syms.some((s: any) => s.name === 'Data'), 'Class Data');
    assert.ok(syms.some((s: any) => s.name === 'Process'), 'Function Process');
    console.log('[PASS] Mixed declarations all extracted');
}

console.log('\n✅ LSP SymbolProvider: 全テスト通過');
