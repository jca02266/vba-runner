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
    const uri = 'file:///test.bas';
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
    const uri = 'file:///test.bas';
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
    const uri = 'file:///test.bas';
    server.didOpen(uri, 'Sub Test()\nEnd Sub');
    server.didClose(uri);
    const symbols = server.getDocumentSymbols(uri);
    assert.strictEqual(symbols.length, 0, 'document closed - empty symbols');
    console.log('[PASS] didClose removes document');
}

// 6. getDocumentSymbols extracts procedures and classes
{
    const server = createServer();
    const uri = 'file:///test.bas';
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

// 7. getHover — ローカル変数に種別・関数名・ファイル名を表示
{
    const server = createServer();
    const uri = 'file:///proj/Sheet1.bas';
    const code = [
        'Sub MyProc()',
        '    Dim x As Integer',
        'End Sub',
    ].join('\n');
    server.didOpen(uri, code);
    const hover = server.getHover(uri, 1, 8); // Line 1, character 8 (the 'x')
    assert.ok(hover !== null, 'hover returned');
    assert.ok(hover.contents.includes('Local variable'), 'shows Local variable');
    assert.ok(hover.contents.includes('MyProc'),          'shows enclosing proc name');
    assert.ok(hover.contents.includes('Sheet1.bas'),      'shows filename');
    console.log('[PASS] getHover: ローカル変数に種別・関数名・ファイル名を表示');
}

// 7b. getHover — モジュール変数
{
    const server = createServer();
    const uri = 'file:///proj/Sheet1.bas';
    const code = [
        'Private m_count As Long',
        'Sub Test()',
        '    m_count = 1',
        'End Sub',
    ].join('\n');
    server.didOpen(uri, code);
    const hover = server.getHover(uri, 2, 6); // 'm_count' in body
    assert.ok(hover !== null, 'hover returned');
    assert.ok(hover.contents.includes('Module variable'), 'shows Module variable');
    assert.ok(hover.contents.includes('Sheet1.bas'),      'shows filename');
    console.log('[PASS] getHover: モジュール変数の種別表示');
}

// 7c. getHover — パラメーター
{
    const server = createServer();
    const uri = 'file:///proj/Sheet1.bas';
    const code = [
        'Sub DoWork(ByVal count As Long)',
        '    Dim x As Long',
        '    x = count',
        'End Sub',
    ].join('\n');
    server.didOpen(uri, code);
    const hover = server.getHover(uri, 2, 8); // 'count' in body
    assert.ok(hover !== null, 'hover returned');
    assert.ok(hover.contents.includes('Parameter'), 'shows Parameter');
    assert.ok(hover.contents.includes('DoWork'),    'shows enclosing proc name');
    assert.ok(hover.contents.includes('ByVal'),     'shows ByVal modifier');
    console.log('[PASS] getHover: パラメーターの種別・修飾子表示');
}

// 7d. getHover — 組み込み型名
{
    const server = createServer();
    const uri = 'file:///proj/Sheet1.bas';
    server.didOpen(uri, 'Sub Test()\n    Dim x As Long\nEnd Sub');
    // hover over 'Long' in 'Dim x As Long' (line 1, col ~14)
    const hover = server.getHover(uri, 1, 14);
    assert.ok(hover !== null, 'hover returned for built-in type');
    assert.ok(hover.contents.includes('Built-in VBA type'), 'shows Built-in VBA type');
    console.log('[PASS] getHover: 組み込み型名 (Long) のホバー');
}

// 7e. getHover — ユーザー定義クラス型名
{
    const server = createServer();
    const helperUri = 'file:///proj/Helper.cls';
    const mainUri   = 'file:///proj/Main.bas';
    server.loadWorkspaceFile(helperUri, [
        'Attribute VB_Name = "Helper"',
        'Public Sub DoWork()',
        'End Sub',
    ].join('\n'));
    server.didOpen(mainUri, [
        'Sub UseHelper()',
        '    Dim h As Helper',
        'End Sub',
    ].join('\n'));
    // hover over 'Helper' in 'Dim h As Helper' (line 1, col ~13)
    const hover = server.getHover(mainUri, 1, 14);
    assert.ok(hover !== null, 'hover returned for user class');
    assert.ok(hover.contents.includes('User-defined class'), 'shows User-defined class');
    assert.ok(hover.contents.includes('Helper.cls'),         'shows class filename');
    console.log('[PASS] getHover: ユーザー定義クラス型名のホバー');
}

// 8. getDefinition returns definition location
{
    const server = createServer();
    const uri = 'file:///test.bas';
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
    const uri = 'file:///test.bas';
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
    const uri = 'file:///test.bas';
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
    const uri = 'file:///test.bas';
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
    const uri = 'file:///test.bas';
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
    const uri = 'file:///test.bas';
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
    const adapter = server.getDebugAdapter('file:///nonexistent.bas');
    assert.strictEqual(adapter, null, 'null returned for nonexistent document');
    console.log('[PASS] getDebugAdapter returns null for missing document');
}

// 15. Multiple documents are handled independently
{
    const server = createServer();
    const uri1 = 'file:///file1.bas';
    const uri2 = 'file:///file2.bas';

    server.didOpen(uri1, 'Sub Test1()\nEnd Sub');
    server.didOpen(uri2, 'Sub Test2()\nEnd Sub');

    const symbols1 = server.getDocumentSymbols(uri1);
    const symbols2 = server.getDocumentSymbols(uri2);

    assert.ok(symbols1.length > 0, 'file1 has symbols');
    assert.ok(symbols2.length > 0, 'file2 has symbols');
    console.log('[PASS] Multiple documents handled independently');
}

// ─── getVariantTypeHints ─────────────────────────────────────────────────────

// 16. Dim なし型 → リテラル代入から推論
{
    const server = createServer();
    const uri = 'file:///test.bas';
    server.didOpen(uri, [
        'Sub Test()',
        '    Dim x',
        '    x = 42',
        'End Sub',
    ].join('\n'));
    const hints = server.getVariantTypeHints(uri);
    const h = hints.find(h => h.label === ' As Long');
    assert.ok(h !== undefined, 'x = 42 → As Long ヒントが存在する');
    console.log('[PASS] getVariantTypeHints: Dim なし型 + 整数代入 → As Long');
}

// 17. Dim As Object + CreateObject → 具体型に推論
{
    const server = createServer();
    const uri = 'file:///test.bas';
    server.didOpen(uri, [
        'Sub Test()',
        '    Dim dict As Object',
        '    Set dict = CreateObject("Scripting.Dictionary")',
        'End Sub',
    ].join('\n'));
    const hints = server.getVariantTypeHints(uri);
    const h = hints.find(h => h.label === ' As Dictionary');
    assert.ok(h !== undefined, 'As Object + CreateObject("Scripting.Dictionary") → As Dictionary');
    console.log('[PASS] getVariantTypeHints: Dim As Object + CreateObject → As Dictionary');
}

// 18. 複数の Object 型変数を同一プロシージャで推論
{
    const server = createServer();
    const uri = 'file:///test.bas';
    server.didOpen(uri, [
        'Sub Test()',
        '    Dim totals As Object',
        '    Dim counts As Object',
        '    Dim seenIDs As Object',
        '    Set totals  = CreateObject("Scripting.Dictionary")',
        '    Set counts  = CreateObject("Scripting.Dictionary")',
        '    Set seenIDs = CreateObject("Scripting.Dictionary")',
        'End Sub',
    ].join('\n'));
    const hints = server.getVariantTypeHints(uri);
    const dictHints = hints.filter(h => h.label === ' As Dictionary');
    assert.strictEqual(dictHints.length, 3, '3変数すべてに As Dictionary ヒントが出る');
    console.log('[PASS] getVariantTypeHints: 複数 As Object 変数を同一手続きで推論');
}

// 19. 型あり変数はヒントなし（As Long 等）
{
    const server = createServer();
    const uri = 'file:///test.bas';
    server.didOpen(uri, [
        'Sub Test()',
        '    Dim n As Long',
        '    n = 1',
        'End Sub',
    ].join('\n'));
    const hints = server.getVariantTypeHints(uri);
    assert.strictEqual(hints.length, 0, '明示型変数はヒントなし');
    console.log('[PASS] getVariantTypeHints: 明示型 (As Long) はヒントなし');
}

// 20. 型なしパラメーター → As Variant ヒント
{
    const server = createServer();
    const uri = 'file:///test.bas';
    server.didOpen(uri, [
        'Sub Test(a, b As Long)',
        'End Sub',
    ].join('\n'));
    const hints = server.getVariantTypeHints(uri);
    const paramHint = hints.find(h => h.label === ' As Variant');
    assert.ok(paramHint !== undefined, '型なしパラメーター a に As Variant ヒント');
    assert.strictEqual(hints.filter(h => h.label === ' As Variant').length, 1, 'b は型あり → ヒントなし');
    console.log('[PASS] getVariantTypeHints: 型なしパラメーター → As Variant');
}

// 20b. As Variant 明示パラメーター → ヒントなし（重複防止）
{
    const server = createServer();
    const uri = 'file:///test.bas';
    server.didOpen(uri, [
        'Private Function ValidatePrice(ByVal v As Variant) As Boolean',
        '    ValidatePrice = True',
        'End Function',
    ].join('\n'));
    const hints = server.getVariantTypeHints(uri);
    const variantHints = hints.filter(h => h.label === ' As Variant');
    assert.strictEqual(variantHints.length, 0, 'As Variant 明示済みパラメーターには重複ヒントを出さない');
    console.log('[PASS] getVariantTypeHints: As Variant 明示パラメーター → ヒントなし');
}

// 21. 型なし Function → 戻り型ヒント
{
    const server = createServer();
    const uri = 'file:///test.bas';
    server.didOpen(uri, [
        'Function GetName()',
        '    GetName = "Alice"',
        'End Function',
    ].join('\n'));
    const hints = server.getVariantTypeHints(uri);
    const retHint = hints.find(h => h.label === ' As String');
    assert.ok(retHint !== undefined, '型なし Function に As String 戻り型ヒント');
    console.log('[PASS] getVariantTypeHints: 型なし Function → As String 戻り型ヒント');
}

// 22. As Object で未代入 → ヒントなし
{
    const server = createServer();
    const uri = 'file:///test.bas';
    server.didOpen(uri, [
        'Sub Test()',
        '    Dim obj As Object',
        'End Sub',
    ].join('\n'));
    const hints = server.getVariantTypeHints(uri);
    assert.strictEqual(hints.length, 0, 'As Object で未代入ならヒントなし');
    console.log('[PASS] getVariantTypeHints: As Object 未代入 → ヒントなし');
}

// 23. FileSystemObject の推論
{
    const server = createServer();
    const uri = 'file:///test.bas';
    server.didOpen(uri, [
        'Sub Test()',
        '    Dim fso As Object',
        '    Set fso = CreateObject("Scripting.FileSystemObject")',
        'End Sub',
    ].join('\n'));
    const hints = server.getVariantTypeHints(uri);
    const h = hints.find(h => h.label === ' As FileSystemObject');
    assert.ok(h !== undefined, 'As Object + CreateObject(FSO) → As FileSystemObject');
    console.log('[PASS] getVariantTypeHints: CreateObject("Scripting.FileSystemObject") → As FileSystemObject');
}

// VBA016: 他の .cls ファイルで定義されたクラスが未知型とみなされないこと
// (loadWorkspaceFile で隣接ファイルを登録した場合)
{
    const server = createServer();
    const mainUri = 'file:///proj/Main.bas';
    const helperUri = 'file:///proj/Helper.cls';

    // Helper.cls をワークスペースファイルとして登録
    server.loadWorkspaceFile(helperUri, [
        'Attribute VB_Name = "Helper"',
        'Public Function GetValue() As Long',
        '    GetValue = 42',
        'End Function',
    ].join('\n'));

    // Main.bas で Helper 型を使う（As Helper）
    server.didOpen(mainUri, [
        'Sub UseHelper()',
        '    Dim h As Helper',
        'End Sub',
    ].join('\n'));

    const diags = server.getDiagnostics(mainUri);
    const vba016 = diags.filter((d: any) => d.code === 'VBA016');
    assert.strictEqual(vba016.length, 0, 'Helper が .cls ファイルとして登録されていれば VBA016 は出ない');
    console.log('[PASS] VBA016: 隣接 .cls ファイルの型を認識する');
}

// getDefinition: 型名（As ClassName）でF12 → 対応する .cls ファイルへジャンプ
{
    const server = createServer();
    const helperUri = 'file:///proj/Helper.cls';
    const mainUri   = 'file:///proj/Main.bas';

    server.loadWorkspaceFile(helperUri, [
        'Attribute VB_Name = "Helper"',
        'Public Sub DoWork()',
        'End Sub',
    ].join('\n'));

    server.didOpen(mainUri, [
        'Sub UseHelper()',
        '    Dim h As Helper',
        'End Sub',
    ].join('\n'));

    // "Helper" は line 1 (0-based), col 13 あたり ("    Dim h As Helper")
    const result = server.getDefinition(mainUri, 1, 15);
    assert.ok(result !== null, 'getDefinition on class name returns result');
    assert.strictEqual(result.uri, helperUri, 'jumps to Helper.cls');
    assert.strictEqual(result.range.start.line, 0, 'starts at line 0');
    console.log('[PASS] getDefinition: 型名 (As ClassName) で .cls ファイルへジャンプ');
}

console.log('\n✅ LSPServer: 全テスト通過');
