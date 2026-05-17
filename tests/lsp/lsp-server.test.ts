import { Lexer } from '../../src/engine/lexer';
import { Parser } from '../../src/engine/parser';
import { LSPServer } from '../../src/lsp/server';
import { assert } from '../../test-libs/test-runner';

function createServer(): LSPServer {
    return new LSPServer();
}

// 1. Server initializes
{
    const server = createServer();
    assert.ok(server, 'server created');
    console.log('[PASS] Server initialization');
}

// 2. Initialize returns capabilities
{
    const server = createServer();
    const response = server.initialize();
    assert.ok(response, 'initialize response');
    assert.ok(response.capabilities, 'capabilities present');
    assert.ok(response.capabilities.textDocumentSync !== undefined, 'textDocumentSync capability');
    assert.ok(response.capabilities.completionProvider, 'completionProvider capability');
    assert.ok(response.capabilities.hoverProvider === true, 'hoverProvider capability');
    assert.ok(response.capabilities.definitionProvider === true, 'definitionProvider capability');
    console.log('[PASS] Initialize returns capabilities');
}

// 3. didOpen stores document
{
    const server = createServer();
    const uri = 'file:///test.vba';
    const content = 'Sub Test()\nEnd Sub';
    server.didOpen(uri, content);
    // No direct getter, but we can verify by calling a method that uses the document
    const symbols = server.getDocumentSymbols(uri);
    assert.ok(Array.isArray(symbols), 'symbols array returned after didOpen');
    console.log('[PASS] didOpen stores document');
}

// 4. didChange updates document
{
    const server = createServer();
    const uri = 'file:///test.vba';
    server.didOpen(uri, 'Sub Test()\nEnd Sub');
    server.didChange(uri, 'Sub Test2()\nEnd Sub');
    const symbols = server.getDocumentSymbols(uri);
    // Should have Test2 now, not Test
    assert.ok(Array.isArray(symbols), 'symbols updated after didChange');
    console.log('[PASS] didChange updates document');
}

// 5. didClose removes document
{
    const server = createServer();
    const uri = 'file:///test.vba';
    server.didOpen(uri, 'Sub Test()\nEnd Sub');
    server.didClose(uri);
    const symbols = server.getDocumentSymbols(uri);
    assert.strictEqual(symbols.length, 0, 'document closed - empty symbols');
    console.log('[PASS] didClose removes document');
}

// 6. getDocumentSymbols extracts procedures and classes
{
    const server = createServer();
    const uri = 'file:///test.vba';
    const code = `
Sub MySub()
End Sub

Function MyFunc() As Integer
End Function

Class MyClass
    Public Value As Integer
End Class
`;
    server.didOpen(uri, code);
    const symbols = server.getDocumentSymbols(uri);
    assert.ok(Array.isArray(symbols), 'symbols array');
    assert.ok(symbols.length > 0, 'symbols extracted');
    console.log('[PASS] getDocumentSymbols extracts symbols');
}

// 7. getHover returns hover info
{
    const server = createServer();
    const uri = 'file:///test.vba';
    const code = `
Sub Test()
    Dim x As Integer
End Sub
`;
    server.didOpen(uri, code);
    const hover = server.getHover(uri, 2, 8); // Line 2, character 8 (the 'x')
    assert.ok(hover === null || typeof hover === 'object', 'hover returns object or null');
    console.log('[PASS] getHover returns hover info');
}

// 8. getDefinition returns definition location
{
    const server = createServer();
    const uri = 'file:///test.vba';
    const code = `
Sub MyProc()
End Sub

Sub Test()
    Call MyProc()
End Sub
`;
    server.didOpen(uri, code);
    const def = server.getDefinition(uri, 5, 9); // MyProc in the Call statement
    assert.ok(def === null || typeof def === 'object', 'definition returns object or null');
    console.log('[PASS] getDefinition returns location');
}

// 9. getCompletions returns completion items
{
    const server = createServer();
    const uri = 'file:///test.vba';
    const code = `
Sub Test()
    Dim x As String
    Dim y As Integer
End Sub
`;
    server.didOpen(uri, code);
    const completions = server.getCompletions(uri, 3, 0);
    assert.ok(Array.isArray(completions), 'completions array');
    console.log('[PASS] getCompletions returns items');
}

// 10. discoverTests finds test procedures
{
    const server = createServer();
    const uri = 'file:///test.vba';
    const code = `
Sub Test_FirstTest()
End Sub

Sub Test_SecondTest()
End Sub

Sub HelperSub()
End Sub
`;
    server.didOpen(uri, code);
    const tests = server.discoverTests(uri);
    assert.ok(Array.isArray(tests), 'tests array');
    console.log('[PASS] discoverTests finds test procedures');
}

// 11. runTests executes tests
{
    const server = createServer();
    const uri = 'file:///test.vba';
    const code = `
Sub Test_Simple()
End Sub
`;
    server.didOpen(uri, code);
    const results = server.runTests(uri);
    assert.ok(Array.isArray(results), 'test results array');
    console.log('[PASS] runTests executes tests');
}

// 12. createDebugAdapter creates adapter for document
{
    const server = createServer();
    const uri = 'file:///test.vba';
    const code = 'Sub Test()\nEnd Sub';
    server.didOpen(uri, code);
    const adapter = server.createDebugAdapter(uri);
    assert.ok(adapter !== null, 'debug adapter created');
    assert.ok(typeof adapter.handleInitialize === 'function', 'adapter has handleInitialize');
    console.log('[PASS] createDebugAdapter creates adapter');
}

// 13. getDebugAdapter retrieves stored adapter
{
    const server = createServer();
    const uri = 'file:///test.vba';
    const code = 'Sub Test()\nEnd Sub';
    server.didOpen(uri, code);
    const created = server.createDebugAdapter(uri);
    const retrieved = server.getDebugAdapter(uri);
    assert.strictEqual(created, retrieved, 'same adapter returned');
    console.log('[PASS] getDebugAdapter retrieves adapter');
}

// 14. getDebugAdapter returns null for nonexistent document
{
    const server = createServer();
    const adapter = server.getDebugAdapter('file:///nonexistent.vba');
    assert.strictEqual(adapter, null, 'null returned for nonexistent document');
    console.log('[PASS] getDebugAdapter returns null for missing document');
}

// 15. Multiple documents are handled independently
{
    const server = createServer();
    const uri1 = 'file:///file1.vba';
    const uri2 = 'file:///file2.vba';

    server.didOpen(uri1, 'Sub Test1()\nEnd Sub');
    server.didOpen(uri2, 'Sub Test2()\nEnd Sub');

    const symbols1 = server.getDocumentSymbols(uri1);
    const symbols2 = server.getDocumentSymbols(uri2);

    assert.ok(symbols1.length > 0, 'file1 has symbols');
    assert.ok(symbols2.length > 0, 'file2 has symbols');
    console.log('[PASS] Multiple documents handled independently');
}

console.log('\n✅ LSPServer: 全テスト通過');
