import * as vscode from 'vscode';
import { Lexer } from './compiler/lexer';
import { Parser } from './compiler/parser';
import { LSPServer } from './lsp/server';

let lspServer: LSPServer;
const documentMap = new Map<string, vscode.TextDocument>();

export async function activate(context: vscode.ExtensionContext) {
    console.log('🚀 VBA Compiler extension activated');
    const outputChannel = vscode.window.createOutputChannel('VBA Compiler');
    outputChannel.appendLine('Extension initialization started...');

    lspServer = new LSPServer();
    outputChannel.appendLine('LSP Server initialized');

    // Register already-open documents
    for (const doc of vscode.workspace.textDocuments) {
        if (doc.languageId === 'vba') {
            documentMap.set(doc.uri.toString(), doc);
            lspServer.didOpen(doc.uri.toString(), doc.getText());
        }
    }

    // Register document open listener
    context.subscriptions.push(
        vscode.workspace.onDidOpenTextDocument((doc) => {
            if (doc.languageId === 'vba') {
                documentMap.set(doc.uri.toString(), doc);
                lspServer.didOpen(doc.uri.toString(), doc.getText());
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
            }
        })
    );

    // Register document close listener
    context.subscriptions.push(
        vscode.workspace.onDidCloseTextDocument((doc) => {
            if (doc.languageId === 'vba') {
                documentMap.delete(doc.uri.toString());
                lspServer.didClose(doc.uri.toString());
            }
        })
    );

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
    const testController = vscode.tests.createTestController('vbaTest', 'VBA Tests');
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

    testController.runHandler = async (request, cancellation) => {
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
                            const message = new vscode.TestMessage(result.error || 'Test failed');
                            test.error = message;
                            test.busy = false;
                        }
                    }
                }
            }
        }
    };

    outputChannel.appendLine('✓ All providers registered successfully');
    outputChannel.appendLine('📝 Open a .vba file and hover over code to test LSP features');
    outputChannel.show();
    console.log('🚀 VBA extension fully initialized');
}

export function deactivate() {
    console.log('VBA Compiler extension deactivated');
}
