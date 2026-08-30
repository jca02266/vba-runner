import { Lexer } from '../../src/engine/lexer';
import { Parser } from '../../src/engine/parser';
import { CompletionProvider } from '../../src/lsp/completion-provider';
import { DefinitionProvider } from '../../src/lsp/definition-provider';
import { buildScopedSymbolTable, lookupSymbolWithContext } from '../../src/lsp/symbol-table';
import { LSPServer } from '../../src/lsp/server';
import { assert } from '../../test-libs/test-runner';

function parse(source: string): any {
    return new Parser(new Lexer(source).tokenize(), { errorRecovery: true }).parse();
}

function positionOf(source: string, needle: string): { line: number; character: number } {
    const lines = source.split('\n');
    const line = lines.findIndex(text => text.includes(needle));
    if (line < 0) throw new Error(`Missing ${needle}`);
    return { line, character: lines[line].indexOf(needle) };
}

const classSource = String.raw`Class PrivateThing
    Private hiddenValue As Long
    Public visibleValue As Long
    Public Sub SetHidden()
        hiddenValue = 1
    End Sub
End Class`;

// Class members are kept in class scope and do not become module/workspace symbols.
{
    const ast = parse(classSource);
    const table = buildScopedSymbolTable(ast.body);
    assert.ok(!table.moduleSymbols.has('hiddenvalue'), 'Private class field is not a module symbol');
    assert.ok(!table.moduleSymbols.has('sethidden'), 'Class procedure is not a module symbol');
    assert.ok(table.classes[0].memberSymbols.has('hiddenvalue'), 'Private field remains in class scope');

    const inside = lookupSymbolWithContext('hiddenValue', positionOf(classSource, 'hiddenValue =').line, table);
    assert.ok(inside?.className === 'PrivateThing', 'Private field resolves inside its class');
    console.log('[PASS] Class Private members stay in class scope');
}

// Workspace completion must not flatten a class from another file into unqualified names.
{
    const mainSource = String.raw`Sub Main()
    hid
End Sub`;
    const provider = new CompletionProvider();
    const completions = provider.getCompletions(
        [...parse(mainSource).body, ...parse(classSource).body],
        mainSource,
        1,
        7,
        parse(mainSource).body,
    );
    assert.ok(!completions.some(item => item.label === 'hiddenValue'), 'Private class field is absent from workspace completion');
    assert.ok(!completions.some(item => item.label === 'SetHidden'), 'Class procedure is absent from workspace completion');
    console.log('[PASS] Workspace completion hides class Private members');
}

// Server definition/references must not resolve a class Private member from another file.
{
    const server = new LSPServer();
    const classUri = 'file:///proj/PrivateThing.bas';
    const mainUri = 'file:///proj/Main.bas';
    const mainSource = String.raw`Sub Main()
    hiddenValue = 2
End Sub`;
    server.didOpen(classUri, classSource);
    server.didOpen(mainUri, mainSource);

    const position = positionOf(mainSource, 'hiddenValue');
    assert.strictEqual(server.getDefinition(mainUri, position.line, position.character), null,
        'Private class member has no definition outside its class');
    assert.strictEqual(server.getReferences(mainUri, position.line, position.character, true).length, 0,
        'Private class member has no references outside its class');
    console.log('[PASS] Server blocks cross-file Private class definition/references');
}

// The member remains available to its own class implementation.
{
    const ast = parse(classSource);
    const provider = new DefinitionProvider();
    provider.setDocumentUri('file:///proj/PrivateThing.bas');
    const position = positionOf(classSource, 'hiddenValue =');
    const definition = provider.getDefinition(ast.body, classSource, position.line, position.character);
    assert.ok(definition, 'Private member definition resolves in class implementation');
    console.log('[PASS] Private class member definition works inside class');
}

console.log('\n✅ LSP Private scope: all tests passed');
