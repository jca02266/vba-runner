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

    // Register references provider (Shift+F12)
    context.subscriptions.push(
        vscode.languages.registerReferenceProvider('vba', {
            provideReferences(document, position, context) {
                const refs = lspServer.getReferences(
                    document.uri.toString(),
                    position.line,
                    position.character,
                    context.includeDeclaration,
                );
                return refs.map((r: any) => new vscode.Location(
                    vscode.Uri.parse(r.uri),
                    new vscode.Range(r.range.start.line, r.range.start.character, r.range.end.line, r.range.end.character),
                ));
            }
        })
    );
    outputChannel.appendLine('✓ References provider registered');

    // Register rename provider (F2)
    context.subscriptions.push(
        vscode.languages.registerRenameProvider('vba', {
            provideRenameEdits(document, position, newName) {
                const edits = lspServer.getRename(
                    document.uri.toString(),
                    position.line,
                    position.character,
                    newName,
                );
                if (!edits) return null;

                const wsEdit = new vscode.WorkspaceEdit();
                for (const edit of edits) {
                    wsEdit.replace(
                        document.uri,
                        new vscode.Range(edit.range.start.line, edit.range.start.character, edit.range.end.line, edit.range.end.character),
                        edit.newText,
                    );
                }
                return wsEdit;
            }
        })
    );
    outputChannel.appendLine('✓ Rename provider registered');

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

    // Register code lens provider
    context.subscriptions.push(
        vscode.languages.registerCodeLensProvider('vba', {
            provideCodeLenses(document) {
                const items = lspServer.getCodeLens(document.uri.toString());
                return items.map((item: any) => {
                    const range = new vscode.Range(
                        item.range.start.line, item.range.start.character,
                        item.range.end.line,   item.range.end.character,
                    );
                    const lens = new vscode.CodeLens(range);
                    lens.command = {
                        title:     item.command.title,
                        command:   item.command.command,
                        arguments: item.command.arguments,
                    };
                    return lens;
                });
            }
        })
    );
    outputChannel.appendLine('✓ Code Lens provider registered');

    // vba-runner.runProcedure: parameterless Sub/Function を即実行
    context.subscriptions.push(
        vscode.commands.registerCommand('vba-runner.runProcedure', async (uri: string, procName: string) => {
            try {
                const { Lexer } = await import('./engine/lexer');
                const { Parser } = await import('./engine/parser');
                const { Evaluator } = await import('./engine/evaluator');
                const doc = documentMap.get(uri);
                if (!doc) return;
                const tokens = new Lexer(doc.getText()).tokenize();
                const ast = new Parser(tokens).parse();
                const ev = new Evaluator((msg: string) => outputChannel.appendLine(msg));
                ev.evaluate(ast);
                const result = ev.callProcedure(procName, []);
                outputChannel.appendLine(`[Run] ${procName}() → ${result}`);
                outputChannel.show();
            } catch (e: any) {
                outputChannel.appendLine(`[Error] ${procName}: ${e.message}`);
                outputChannel.show();
            }
        })
    );

    // vba-runner.findReferences: CodeLens から呼ばれる「参照を検索」コマンド
    // uri, line, character は CodeLens が宣言位置として渡す
    context.subscriptions.push(
        vscode.commands.registerCommand('vba-runner.findReferences', async (uri: string, line: number, character: number) => {
            const docUri = vscode.Uri.parse(uri);
            const pos = new vscode.Position(line, character);
            await vscode.window.showTextDocument(docUri, { selection: new vscode.Range(pos, pos) });
            vscode.commands.executeCommand('editor.action.referenceSearch.trigger');
        })
    );

    // vba-runner.generateTest / vba-runner.goToTest: stub（将来実装）
    context.subscriptions.push(
        vscode.commands.registerCommand('vba-runner.generateTest', () => {
            vscode.window.showInformationMessage('Test generation: coming soon');
        })
    );
    context.subscriptions.push(
        vscode.commands.registerCommand('vba-runner.goToTest', () => {
            vscode.commands.executeCommand('workbench.action.findInFiles');
        })
    );

    outputChannel.appendLine('✓ All providers registered successfully');
    outputChannel.appendLine('📝 Open a .bas file and hover over code to test LSP features');
    outputChannel.show();
    console.log('🚀 VBA extension fully initialized');
}

export function deactivate() {
    console.log('VBA Runner extension deactivated');
}
