import * as vscode from 'vscode';
import { LSPServer } from './lsp/server';
import { VBADebugAdapterFactory } from './lsp/vscode-debug-adapter';

let lspServer: LSPServer;
const documentMap = new Map<string, vscode.TextDocument>();

export async function activate(context: vscode.ExtensionContext) {
    console.log('🚀 VBA Runner extension activated');
    const outputChannel = vscode.window.createOutputChannel('VBA Runner');
    outputChannel.appendLine('Extension initialization started...');

    lspServer = new LSPServer();
    outputChannel.appendLine('LSP Server initialized');

    // Create diagnostic collection for VBA parse errors
    const diagnosticCollection = vscode.languages.createDiagnosticCollection('vba');
    context.subscriptions.push(diagnosticCollection);

    function updateDiagnostics(uri: vscode.Uri): void {
        const raw = lspServer.getDiagnostics(uri.toString());
        const diags = raw.map((d: any) => {
            const range = new vscode.Range(
                d.range.start.line,
                d.range.start.character,
                d.range.end.line,
                d.range.end.character
            );
            const sev = d.severity === 1 ? vscode.DiagnosticSeverity.Error : vscode.DiagnosticSeverity.Warning;
            const diag = new vscode.Diagnostic(range, d.message, sev);
            diag.source = d.source;
            return diag;
        });
        diagnosticCollection.set(uri, diags);
    }

    // Register already-open documents
    for (const doc of vscode.workspace.textDocuments) {
        if (doc.languageId === 'vba') {
            documentMap.set(doc.uri.toString(), doc);
            lspServer.didOpen(doc.uri.toString(), doc.getText());
            updateDiagnostics(doc.uri);
        }
    }

    // Register document open listener
    context.subscriptions.push(
        vscode.workspace.onDidOpenTextDocument((doc) => {
            if (doc.languageId === 'vba') {
                documentMap.set(doc.uri.toString(), doc);
                lspServer.didOpen(doc.uri.toString(), doc.getText());
                updateDiagnostics(doc.uri);
            }
        })
    );

    // Register document change listener
    context.subscriptions.push(
        vscode.workspace.onDidChangeTextDocument((event) => {
            const doc = event.document;
            if (doc.languageId === 'vba') {
                documentMap.set(doc.uri.toString(), doc);
                lspServer.didChange(doc.uri.toString(), doc.getText());
                updateDiagnostics(doc.uri);
            }
        })
    );

    // Register document close listener
    context.subscriptions.push(
        vscode.workspace.onDidCloseTextDocument((doc) => {
            if (doc.languageId === 'vba') {
                documentMap.delete(doc.uri.toString());
                lspServer.didClose(doc.uri.toString());
                diagnosticCollection.delete(doc.uri);
            }
        })
    );
    outputChannel.appendLine('✓ Diagnostics collection registered');

    // Register hover provider
    context.subscriptions.push(
        vscode.languages.registerHoverProvider('vba', {
            provideHover(document, position) {
                const hover = lspServer.getHover(
                    document.uri.toString(),
                    position.line,
                    position.character
                );
                if (!hover) return null;

                const markdownString = new vscode.MarkdownString();
                markdownString.appendMarkdown(`\`\`\`vba\n${hover.contents}\n\`\`\``);

                return new vscode.Hover(markdownString, hover.range ?
                    new vscode.Range(
                        hover.range.start.line,
                        hover.range.start.character,
                        hover.range.end.line,
                        hover.range.end.character
                    ) : undefined
                );
            }
        })
    );
    outputChannel.appendLine('✓ Hover provider registered');

    // Register definition provider
    context.subscriptions.push(
        vscode.languages.registerDefinitionProvider('vba', {
            provideDefinition(document, position) {
                const def = lspServer.getDefinition(
                    document.uri.toString(),
                    position.line,
                    position.character
                );
                if (!def) return null;

                return new vscode.Location(
                    vscode.Uri.parse(def.uri),
                    new vscode.Range(
                        def.range.start.line,
                        def.range.start.character,
                        def.range.end.line,
                        def.range.end.character
                    )
                );
            }
        })
    );
    outputChannel.appendLine('✓ Definition provider registered');

    // Register document symbol provider (outline)
    function toVscodeSymbol(sym: any): vscode.DocumentSymbol {
        const range = new vscode.Range(
            sym.location.range.start.line,
            sym.location.range.start.character,
            sym.location.range.end.line,
            sym.location.range.end.character
        );
        // SymbolProvider enum is 1-based; vscode.SymbolKind is 0-based
        const kind = (sym.kind - 1) as vscode.SymbolKind;
        const ds = new vscode.DocumentSymbol(sym.name, sym.detail ?? '', kind, range, range);
        if (sym.children?.length) {
            ds.children = sym.children.map(toVscodeSymbol);
        }
        return ds;
    }

    context.subscriptions.push(
        vscode.languages.registerDocumentSymbolProvider('vba', {
            provideDocumentSymbols(document) {
                const symbols = lspServer.getDocumentSymbols(document.uri.toString());
                return symbols.map(toVscodeSymbol);
            }
        })
    );
    outputChannel.appendLine('✓ DocumentSymbol provider registered');

    // Register completion provider
    context.subscriptions.push(
        vscode.languages.registerCompletionItemProvider('vba', {
            provideCompletionItems(document, position) {
                const completions = lspServer.getCompletions(
                    document.uri.toString(),
                    position.line,
                    position.character
                );

                return completions.map((item: any) => {
                    const ci = new vscode.CompletionItem(item.label, item.kind || vscode.CompletionItemKind.Function);
                    ci.detail = item.detail;
                    ci.documentation = item.documentation;
                    return ci;
                });
            }
        }, '.')); // Trigger on dot
    outputChannel.appendLine('✓ Completion provider registered');

    // Register test discovery
    const testController = vscode.tests.createTestController('vbaRunner', 'VBA Tests');
    context.subscriptions.push(testController);

    const createTestItems = (document: vscode.TextDocument) => {
        const tests = lspServer.discoverTests(document.uri.toString());

        for (const test of tests) {
            const range = new vscode.Range(
                test.range.start.line,
                test.range.start.character,
                test.range.end.line,
                test.range.end.character
            );

            const item = testController.createTestItem(
                test.id,
                test.label,
                document.uri
            );
            item.range = range;
            testController.items.add(item);
        }
    };

    // Create test items for open documents
    for (const doc of vscode.workspace.textDocuments) {
        if (doc.languageId === 'vba') {
            createTestItems(doc);
        }
    }

    // Register test runner
    testController.resolveHandler = async (item) => {
        if (!item) {
            // Discover all tests
            for (const doc of vscode.workspace.textDocuments) {
                if (doc.languageId === 'vba') {
                    createTestItems(doc);
                }
            }
        }
    };

    testController.createRunProfile('Run', vscode.TestRunProfileKind.Run, async (request, _cancellation) => {
        for (const item of request.include || []) {
            const uri = item.uri?.toString();
            if (uri) {
                const results = lspServer.runTests(uri);

                for (const result of results) {
                    const test = testController.items.get(result.id);
                    if (test) {
                        if (result.state === 'passed') {
                            test.busy = false;
                        } else if (result.state === 'failed') {
                            test.error = result.error || 'Test failed';
                            test.busy = false;
                        }
                    }
                }
            }
        }
    }, true);

    // Register DAP debug adapter factory
    context.subscriptions.push(
        vscode.debug.registerDebugAdapterDescriptorFactory('vba', new VBADebugAdapterFactory(lspServer))
    );
    outputChannel.appendLine('✓ Debug adapter factory registered');

    outputChannel.appendLine('✓ All providers registered successfully');
    outputChannel.appendLine('📝 Open a .bas file and hover over code to test LSP features');
    outputChannel.show();
    console.log('🚀 VBA extension fully initialized');
}

export function deactivate() {
    console.log('VBA Runner extension deactivated');
}
