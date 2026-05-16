import * as vscode from 'vscode';
import { Lexer } from './compiler/lexer';
import { Parser } from './compiler/parser';
import { LSPServer } from './lsp/server';

let lspServer: LSPServer;
const documentMap = new Map<string, vscode.TextDocument>();

/**
 * Web Extension version: Uses vscode.workspace.fs for file operations
 * Compatible with VSCode for the Web (vscode.dev, github.dev)
 */
export async function activate(context: vscode.ExtensionContext) {
    console.log('🚀 VBA Compiler Web Extension activated');
    const outputChannel = vscode.window.createOutputChannel('VBA Compiler');
    outputChannel.appendLine('Web Extension initialization started...');

    lspServer = new LSPServer();
    outputChannel.appendLine('LSP Server initialized');

    // Register already-open documents
    for (const doc of vscode.workspace.textDocuments) {
        if (isVBADocument(doc)) {
            documentMap.set(doc.uri.toString(), doc);
            lspServer.didOpen(doc.uri.toString(), doc.getText());
        }
    }

    // Register document open listener
    context.subscriptions.push(
        vscode.workspace.onDidOpenTextDocument((doc) => {
            if (isVBADocument(doc)) {
                documentMap.set(doc.uri.toString(), doc);
                lspServer.didOpen(doc.uri.toString(), doc.getText());
            }
        })
    );

    // Register document change listener
    context.subscriptions.push(
        vscode.workspace.onDidChangeTextDocument((event) => {
            const doc = event.document;
            if (isVBADocument(doc)) {
                documentMap.set(doc.uri.toString(), doc);
                lspServer.didChange(doc.uri.toString(), doc.getText());
            }
        })
    );

    // Register document close listener
    context.subscriptions.push(
        vscode.workspace.onDidCloseTextDocument((doc) => {
            if (isVBADocument(doc)) {
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
        }, '.')
    );

    // Register test discovery and execution
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
        if (isVBADocument(doc)) {
            createTestItems(doc);
        }
    }

    // File change watcher for test discovery
    const vbaWatcher = vscode.workspace.createFileSystemWatcher('**/*.{vba,cls,bas,frm}');
    context.subscriptions.push(vbaWatcher);

    vbaWatcher.onDidCreate(async (uri) => {
        const doc = await vscode.workspace.openTextDocument(uri);
        if (isVBADocument(doc)) {
            createTestItems(doc);
        }
    });

    // Test resolver and runner
    testController.resolveHandler = async (item) => {
        if (!item) {
            // Discover all tests
            for (const doc of vscode.workspace.textDocuments) {
                if (isVBADocument(doc)) {
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

    // Command: Open Web UI for VBA development
    context.subscriptions.push(
        vscode.commands.registerCommand('vba-compiler.openWebUI', async () => {
            const panel = vscode.window.createWebviewPanel(
                'vbaWebUI',
                'VBA Web Compiler',
                vscode.ViewColumn.Beside,
                {
                    enableScripts: true,
                    localResourceRoots: [context.extensionUri]
                }
            );

            panel.webview.html = getWebUIHTML(panel.webview, context.extensionUri);
        })
    );

    console.log('VBA Web Extension fully initialized');
}

function isVBADocument(doc: vscode.TextDocument): boolean {
    return doc.languageId === 'vba' ||
        doc.uri.fsPath.match(/\.(vba|bas|cls|frm)$/i) !== null;
}

function getWebUIHTML(webview: vscode.Webview, extensionUri: vscode.Uri): string {
    return `<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>VBA Web Compiler</title>
    <style>
        body {
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif;
            display: flex;
            height: 100vh;
            margin: 0;
            background: var(--vscode-editor-background);
            color: var(--vscode-editor-foreground);
        }
        .editor-pane {
            flex: 1;
            display: flex;
            flex-direction: column;
            border-right: 1px solid var(--vscode-panel-border);
        }
        .output-pane {
            flex: 1;
            display: flex;
            flex-direction: column;
            border-left: 1px solid var(--vscode-panel-border);
        }
        textarea, pre {
            font-family: 'Monaco', 'Menlo', 'Ubuntu Mono', monospace;
            font-size: 13px;
        }
        textarea {
            flex: 1;
            padding: 8px;
            border: none;
            background: var(--vscode-editor-background);
            color: var(--vscode-editor-foreground);
        }
        #output {
            flex: 1;
            padding: 8px;
            overflow-y: auto;
            background: var(--vscode-panel-background);
        }
        .toolbar {
            padding: 8px;
            border-bottom: 1px solid var(--vscode-panel-border);
            display: flex;
            gap: 8px;
        }
        button {
            padding: 6px 12px;
            background: var(--vscode-button-background);
            color: var(--vscode-button-foreground);
            border: none;
            border-radius: 2px;
            cursor: pointer;
        }
        button:hover {
            background: var(--vscode-button-hoverBackground);
        }
    </style>
</head>
<body>
    <div class="editor-pane">
        <div class="toolbar">
            <button onclick="runCode()">Run</button>
            <button onclick="clearOutput()">Clear Output</button>
        </div>
        <textarea id="editor" placeholder="Enter VBA code here..."></textarea>
    </div>
    <div class="output-pane">
        <div class="toolbar">Output</div>
        <pre id="output"></pre>
    </div>
    <script>
        const vscode = acquireVsCodeApi();

        function runCode() {
            const code = document.getElementById('editor').value;
            vscode.postMessage({
                command: 'run',
                code: code
            });
        }

        function clearOutput() {
            document.getElementById('output').textContent = '';
        }

        window.addEventListener('message', event => {
            const message = event.data;
            if (message.command === 'output') {
                const output = document.getElementById('output');
                output.textContent += message.text;
                output.scrollTop = output.scrollHeight;
            }
        });
    </script>
</body>
</html>`;
}

export function deactivate() {
    console.log('VBA Web Extension deactivated');
}
