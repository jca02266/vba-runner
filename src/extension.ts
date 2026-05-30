import * as vscode from 'vscode';

import * as path from 'path';
import * as fs from 'fs';
import { LSPServer } from './lsp/server';
import { VBADebugAdapterFactory } from './lsp/vscode-debug-adapter';
import { Lexer } from './engine/lexer';
import { Parser } from './engine/parser';
import { Evaluator } from './engine/evaluator';
import { format as vbaFormat } from './lsp/formatter';
import { FoldingRangeProvider as VBAFoldingRangeProvider } from './lsp/folding-range-provider';
import { generateCallGraphHtml, generateDrawioXml } from './lsp/call-graph-webview';
import { findMatchingExpressions } from './lsp/ast-comparison';
import { needsLineContinuation } from './lsp/line-continuation-checker';

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

    function shouldShowLintDiag(code: string | undefined): boolean {
        const config = vscode.workspace.getConfiguration('vba-runner');
        const enabledCodes = config.get<string[]>('lint.enabledCodes', []);
        if (enabledCodes.length > 0) return enabledCodes.includes(code ?? '');
        return config.get('lint.enabled', false);
    }

    function updateDiagnostics(uri: vscode.Uri): void {
        const raw = lspServer.getDiagnostics(uri.toString());
        const filtered = raw.filter((d: any) => !d.code || shouldShowLintDiag(String(d.code)));
        const diags = filtered.map((d: any) => {
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
                createTestItems(doc);
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

    // Refresh diagnostics when the setting changes
    context.subscriptions.push(
        vscode.workspace.onDidChangeConfiguration((event) => {
            if (event.affectsConfiguration('vba-runner.lint.enabled') || event.affectsConfiguration('vba-runner.lint.enabledCodes')) {
                for (const [uriStr] of documentMap) {
                    updateDiagnostics(vscode.Uri.parse(uriStr));
                }
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
        const run = testController.createTestRun(request);
        const items = request.include?.length ? request.include : [...testController.items].map(([, item]) => item);
        for (const item of items) {
            run.started(item);
            const uri = item.uri?.toString();
            if (uri) {
                const results = lspServer.runTests(uri);
                for (const result of results) {
                    const test = testController.items.get(result.id);
                    if (!test) continue;
                    if (result.state === 'passed') {
                        run.passed(test);
                    } else {
                        run.failed(test, new vscode.TestMessage(result.error ?? 'Test failed'));
                    }
                }
            }
        }
        run.end();
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

    // AssertHelper クラス定義（テスト実行時に自動注入）
    const ASSERT_HELPER_SRC = `Class AssertHelper
Option Explicit
Public Failed As Boolean
Public FailMessage As String
Sub Reset()
    Failed = False
    FailMessage = ""
End Sub
Sub Assert(actual, expected, message)
    If actual <> expected Then
        Debug.Print "[FAIL] " & message
        Debug.Print "  Expected: " & CStr(expected)
        Debug.Print "  Actual  : " & CStr(actual)
        Failed = True
        FailMessage = message
        Err.Raise vbObjectError + 1, "Assert", message
    End If
End Sub
Sub IsTrue(value, message)
    If Not CBool(value) Then
        Debug.Print "[FAIL] " & message & " (expected True)"
        Failed = True
        FailMessage = message
        Err.Raise vbObjectError + 1, "Assert", message
    End If
End Sub
Sub IsFalse(value, message)
    If CBool(value) Then
        Debug.Print "[FAIL] " & message & " (expected False)"
        Failed = True
        FailMessage = message
        Err.Raise vbObjectError + 1, "Assert", message
    End If
End Sub
End Class`;

    // vba-runner.runProcedure: Sub/Function を即実行
    // isTestProc=true のとき: AssertHelper を注入したラッパー Sub を生成して呼び出す
    context.subscriptions.push(
        vscode.commands.registerCommand('vba-runner.runProcedure', (uri: string, procName: string, isTestProc?: boolean) => {
            try {
                const doc = documentMap.get(uri);
                if (!doc) {
                    outputChannel.appendLine(`[Error] Document not found in map: ${uri}`);
                    outputChannel.show();
                    return;
                }

                // 同ディレクトリの .bas/.cls を収集して結合
                const dir = path.dirname(doc.uri.fsPath);
                const entries = fs.readdirSync(dir).filter(f => /\.(bas|cls)$/i.test(f));
                const parts: string[] = [ASSERT_HELPER_SRC];
                for (const entry of entries) {
                    // ディレクトリに AssertHelper.cls があっても重複注入しない
                    if (/^AssertHelper\.cls$/i.test(entry)) continue;
                    const filePath = path.join(dir, entry);
                    const content = fs.readFileSync(filePath, 'utf8');
                    if (/\.cls$/i.test(entry)) {
                        const className = path.basename(entry, path.extname(entry));
                        parts.push(`Class ${className}\n${content}\nEnd Class`);
                    } else {
                        parts.push(content);
                    }
                }

                let callTarget = procName;
                if (isTestProc) {
                    // ラッパー Sub を生成: assert インスタンスを作って渡し、結果を表示
                    const wrapperName = `__VBARunner_${procName}__`;
                    parts.push([
                        `Sub ${wrapperName}()`,
                        `    Dim assert As New AssertHelper`,
                        `    ${procName} assert`,
                        `    If assert.Failed Then`,
                        `        Debug.Print "[FAIL] ${procName}: " & assert.FailMessage`,
                        `    Else`,
                        `        Debug.Print "[PASS] ${procName}"`,
                        `    End If`,
                        `End Sub`,
                    ].join('\n'));
                    callTarget = wrapperName;
                }

                const combined = parts.join('\n\n');
                const tokens = new Lexer(combined).tokenize();
                const ast = new Parser(tokens).parse();
                const ev = new Evaluator((msg: string) => outputChannel.appendLine(msg));
                ev.evaluate(ast);
                const result = ev.callProcedure(callTarget, []);
                if (result !== undefined) {
                    outputChannel.appendLine(`[Run] ${procName}() → ${result}`);
                }
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

    // vba-runner.generateTest: stub（将来実装）
    context.subscriptions.push(
        vscode.commands.registerCommand('vba-runner.generateTest', () => {
            vscode.window.showInformationMessage('Test generation: coming soon');
        })
    );

    // vba-runner.goToTest: シンボル検索で Test_<procName> 関数へジャンプ
    context.subscriptions.push(
        vscode.commands.registerCommand('vba-runner.goToTest', async (uri: string, procName: string) => {
            const testName = `Test_${procName}`;
            const symbols = lspServer.getDocumentSymbols(uri);
            const testSymbol = symbols.find((s: any) =>
                s.name.toLowerCase() === testName.toLowerCase()
            );
            if (testSymbol) {
                const docUri = vscode.Uri.parse(uri);
                const pos = new vscode.Position(
                    testSymbol.location.range.start.line,
                    testSymbol.location.range.start.character
                );
                await vscode.window.showTextDocument(docUri, { selection: new vscode.Range(pos, pos) });
            } else {
                vscode.window.showInformationMessage(`テスト関数 '${testName}' が見つかりません`);
            }
        })
    );

    // Register document formatting provider (Shift+Alt+F)
    context.subscriptions.push(
        vscode.languages.registerDocumentFormattingEditProvider('vba', {
            provideDocumentFormattingEdits(document, formattingOptions) {
                const source = document.getText();
                const edits = vbaFormat(source, {
                    indentSize: formattingOptions.tabSize,
                    indentChar: formattingOptions.insertSpaces ? ' ' : '\t',
                });
                return edits.map(e => new vscode.TextEdit(
                    new vscode.Range(
                        e.range.start.line, e.range.start.character,
                        e.range.end.line, e.range.end.character,
                    ),
                    e.newText,
                ));
            }
        })
    );
    outputChannel.appendLine('✓ Document formatting provider registered');

    // Register folding range provider (enables sticky scroll for Sub/For/While/With/If etc.)
    // Parse directly from the document text to avoid LSP cache timing issues.
    const vbaFoldingProvider = new VBAFoldingRangeProvider();
    context.subscriptions.push(
        vscode.languages.registerFoldingRangeProvider('vba', {
            provideFoldingRanges(document) {
                try {
                    const source = document.getText();
                    const tokens = new Lexer(source).tokenize();
                    const ast = new Parser(tokens).parse();
                    const ranges = vbaFoldingProvider.getFoldingRanges(ast.body);
                    return ranges.map(r => new vscode.FoldingRange(r.startLine, r.endLine));
                } catch {
                    return [];
                }
            }
        })
    );
    outputChannel.appendLine('✓ Folding range provider registered');

    // Register inlay hints provider (loop-continue GoTo + Variant type inference)
    context.subscriptions.push(
        vscode.languages.registerInlayHintsProvider('vba', {
            provideInlayHints(document) {
                const uri = document.uri.toString();
                const gotoHints    = lspServer.getInlayHints(uri);
                const variantHints = lspServer.getVariantTypeHints(uri);

                return [...gotoHints, ...variantHints].map(h => new vscode.InlayHint(
                    new vscode.Position(h.line, h.character),
                    h.label,
                    vscode.InlayHintKind.Type,
                ));
            }
        })
    );
    outputChannel.appendLine('✓ Inlay hints provider registered');

    // Register on-type formatting provider (auto line continuation _)
    context.subscriptions.push(
        vscode.languages.registerOnTypeFormattingEditProvider('vba', {
            provideOnTypeFormattingEdits(document, position) {
                const config = vscode.workspace.getConfiguration('vba-runner');
                if (!config.get('editor.autoLineContinuation', true)) return [];
                if (position.line === 0) return [];

                const prevLine = document.lineAt(position.line - 1);
                const trimmed = prevLine.text.trimEnd();

                if (!needsLineContinuation(trimmed)) return [];

                // Insert ' _' replacing trailing whitespace on the previous line
                const insertPos = new vscode.Position(position.line - 1, trimmed.length);
                const lineEnd   = new vscode.Position(position.line - 1, prevLine.text.length);
                return [vscode.TextEdit.replace(new vscode.Range(insertPos, lineEnd), ' _')];
            }
        }, '\n')
    );
    outputChannel.appendLine('✓ On-type formatting provider (line continuation) registered');

    // Helper function for showing call graph panel
    function showCallGraphPanel(uri: string, focusProcName: string | null): void {
        try {
            // スキャン対象ディレクトリを決定
            const docUri = vscode.Uri.parse(uri);
            const dir = path.dirname(docUri.fsPath);

            // 同ディレクトリの全 .bas/.cls ファイルをスキャン
            const fileContents = new Map<string, string>();
            const entries = fs.readdirSync(dir).filter(f => /\.(bas|cls)$/i.test(f));
            for (const entry of entries) {
                const filePath = path.join(dir, entry);
                try {
                    const content = fs.readFileSync(filePath, 'utf8');
                    const fileUri = vscode.Uri.file(filePath).toString();
                    fileContents.set(fileUri, content);
                } catch (e) {
                    outputChannel.appendLine(`[Warning] Failed to read ${entry}: ${(e as any).message}`);
                }
            }

            // LSP Server で複数ファイルからコールグラフを生成
            const graph = lspServer.buildCallGraphFromFiles(fileContents);

            const panel = vscode.window.createWebviewPanel(
                'vbaCallGraph',
                'VBA Call Graph',
                vscode.ViewColumn.Beside,
                { enableScripts: true, localResourceRoots: [] }
            );
            panel.webview.html = generateCallGraphHtml(graph, panel.webview, context.extensionUri, focusProcName);

            panel.webview.onDidReceiveMessage(
                async (msg: any) => {
                    if (msg.type === 'goToDefinition') {
                        const procName = msg.procName;
                        const node = graph.nodes.get(procName.toLowerCase());
                        if (node) {
                            const docUri = vscode.Uri.parse(node.uri);
                            const pos = new vscode.Position(node.line, 0);
                            await vscode.window.showTextDocument(docUri, { selection: new vscode.Range(pos, pos) });
                        }
                    } else if (msg.type === 'showFullGraph') {
                        showCallGraphPanel(uri, null);
                        panel.dispose();
                    } else if (msg.type === 'saveDrawio') {
                        const xml = generateDrawioXml(graph, focusProcName);
                        const saveUri = await vscode.window.showSaveDialog({
                            filters: { 'Draw.io': ['drawio'] },
                            defaultUri: vscode.Uri.file(path.join(path.dirname(docUri.fsPath), 'call-graph.drawio')),
                        });
                        if (saveUri) {
                            fs.writeFileSync(saveUri.fsPath, xml, 'utf8');
                            vscode.window.showInformationMessage(`Saved: ${saveUri.fsPath}`);
                        }
                    }
                },
                undefined,
                context.subscriptions
            );
        } catch (e: any) {
            vscode.window.showErrorMessage(`Failed to show call graph: ${e.message}`);
            outputChannel.appendLine(`[Error] Call graph: ${e.message}`);
            outputChannel.show();
        }
    }

    // vba-runner.showCallGraph: Show call graph for the current document
    context.subscriptions.push(
        vscode.commands.registerCommand('vba-runner.showCallGraph', () => {
            const editor = vscode.window.activeTextEditor;
            if (!editor || editor.document.languageId !== 'vba') {
                vscode.window.showWarningMessage('Please open a VBA file first');
                return;
            }
            showCallGraphPanel(editor.document.uri.toString(), null);
        })
    );

    // vba-runner.showInCallGraph: Show call graph focused on a procedure (from Code Lens)
    context.subscriptions.push(
        vscode.commands.registerCommand('vba-runner.showInCallGraph', (uri: string, procName: string) => {
            showCallGraphPanel(uri, procName);
        })
    );
    outputChannel.appendLine('✓ Call graph commands registered');

    // Register Code Actions (Refactor menu)
    context.subscriptions.push(
        vscode.languages.registerCodeActionsProvider('vba', {
            provideCodeActions(document, range) {
                const actions: vscode.CodeAction[] = [];

                if (!range.isEmpty) {
                    const action = new vscode.CodeAction('Introduce Variable', vscode.CodeActionKind.Refactor);
                    action.command = {
                        title: 'Introduce Variable',
                        command: 'vba-runner.introduceVariable',
                        arguments: [document.uri, range],
                    };
                    actions.push(action);
                }

                return actions;
            }
        })
    );

    context.subscriptions.push(
        vscode.commands.registerCommand('vba-runner.introduceVariable', async (uri: vscode.Uri, range: vscode.Range) => {
            const editor = vscode.window.activeTextEditor;
            if (!editor || editor.document.uri.toString() !== uri.toString()) {
                vscode.window.showWarningMessage('Document context lost');
                return;
            }

            const selectedText = editor.document.getText(range);
            if (!selectedText.trim()) {
                vscode.window.showWarningMessage('Please select an expression');
                return;
            }

            const varName = await vscode.window.showInputBox({
                prompt: 'Variable name:',
                value: 'result',
                validateInput: (input) => {
                    if (!input) return 'Variable name is required';
                    if (!/^[a-zA-Z_][a-zA-Z0-9_]*$/.test(input)) {
                        return 'Invalid variable name (must start with letter or underscore)';
                    }
                    return '';
                },
            });

            if (!varName) return;

            const insertLine = range.start.line;
            const currentLineText = editor.document.lineAt(insertLine).text;
            const indent = currentLineText.match(/^\s*/)?.[0] || '';

            const insertPosition = new vscode.Position(insertLine, 0);
            const dimAndAssign = `${indent}Dim ${varName}\n${indent}${varName} = ${selectedText}\n`;

            // First, insert Dim and assignment
            const insertEdit = new vscode.WorkspaceEdit();
            insertEdit.insert(uri, insertPosition, dimAndAssign);
            await vscode.workspace.applyEdit(insertEdit);

            // Parse selected expression to AST for comparison
            let selectedExprAst: any = null;
            try {
                const exprTokens = new Lexer(selectedText).tokenize();
                const exprParser = new Parser(exprTokens);
                selectedExprAst = exprParser.parseExpressionPublic();
            } catch (e) {
                // If parsing fails, use null to skip AST comparison
            }

            // Find all matching expressions in the procedure
            const replacementOffsets: Array<{ line: number; start: number; end: number }> = [];

            if (selectedExprAst) {
                const ast = lspServer.parseDocument(editor.document.getText());
                if (ast?.body) {
                    // Find the procedure containing the current line
                    for (const stmt of ast.body) {
                        if (stmt.type === 'ProcedureDeclaration') {
                            const procStart = stmt.loc?.start.line ?? 0;
                            const procEnd = stmt.loc?.end.line ?? 0;
                            if (procStart <= range.start.line + 1 && range.start.line + 1 <= procEnd) {
                                const matches = findMatchingExpressions(stmt.body, selectedExprAst);
                                for (const match of matches) {
                                    replacementOffsets.push({
                                        line: match.start.line,
                                        start: match.start.column,
                                        end: match.end.column
                                    });
                                }
                                break;
                            }
                        }
                    }
                }
            }

            // Remove the original selection from replacements
            const otherReplacements = replacementOffsets.filter(r =>
                !(r.line === range.start.line && r.start === range.start.character && r.end === range.end.character)
            );

            // Note: Don't apply other replacements yet; will ask user after showing multi-cursor

            // Then, replace the original expression (accounting for Dim and assignment added earlier)
            const replacementStart = new vscode.Position(range.start.line + 2, range.start.character);
            const replacementEnd = new vscode.Position(range.end.line + 2, range.end.character);
            const replaceEdit = new vscode.WorkspaceEdit();
            replaceEdit.replace(uri, new vscode.Range(replacementStart, replacementEnd), varName);
            await vscode.workspace.applyEdit(replaceEdit);

            // Set multi-cursor on all occurrences of the variable name
            const dimLineText = editor.document.lineAt(insertLine).text;
            const assignLineText = editor.document.lineAt(insertLine + 1).text;
            const dimVarPos = dimLineText.indexOf(varName);
            const assignVarPos = assignLineText.indexOf(varName);

            const selections: vscode.Selection[] = [];
            if (dimVarPos >= 0) {
                selections.push(new vscode.Selection(
                    insertLine, dimVarPos,
                    insertLine, dimVarPos + varName.length
                ));
            }
            if (assignVarPos >= 0) {
                selections.push(new vscode.Selection(
                    insertLine + 1, assignVarPos,
                    insertLine + 1, assignVarPos + varName.length
                ));
            }

            // Add selections for all matched and replaced expressions
            for (const repl of otherReplacements) {
                selections.push(new vscode.Selection(
                    repl.line, repl.start,
                    repl.line, repl.start + varName.length
                ));
            }

            selections.push(new vscode.Selection(
                range.start.line + 2, range.start.character,
                range.start.line + 2, range.start.character + varName.length
            ));

            editor.selections = selections;
            editor.revealRange(new vscode.Range(selections[0].start, selections[0].end));

            // Ask user about other replacements after multi-cursor is set
            if (otherReplacements.length > 0) {
                const choice = await vscode.window.showQuickPick(
                    ['全て置換', '一つずつ確認', '置換しない'],
                    { placeHolder: `他に ${otherReplacements.length} 箇所見つかりました` }
                );

                if (choice === '全て置換') {
                    // Apply all replacements from back to front to avoid offset changes
                    otherReplacements.sort((a, b) => b.line - a.line || b.start - a.start);
                    for (const repl of otherReplacements) {
                        const replaceEdit = new vscode.WorkspaceEdit();
                        replaceEdit.replace(uri, new vscode.Range(repl.line, repl.start, repl.line, repl.end), varName);
                        await vscode.workspace.applyEdit(replaceEdit);
                    }
                } else if (choice === '一つずつ確認') {
                    // Apply replacements one by one, asking for each
                    otherReplacements.sort((a, b) => b.line - a.line || b.start - a.start);
                    for (const repl of otherReplacements) {
                        const lineText = editor.document.lineAt(repl.line).text;
                        const exprText = lineText.substring(repl.start, repl.end);

                        const userChoice = await vscode.window.showQuickPick(
                            ['置換', 'スキップ'],
                            {
                                placeHolder: `行 ${repl.line + 1}: "${exprText}" を置換しますか？`,
                                canPickMany: false
                            }
                        );

                        if (userChoice === '置換') {
                            const replaceEdit = new vscode.WorkspaceEdit();
                            replaceEdit.replace(uri, new vscode.Range(repl.line, repl.start, repl.line, repl.end), varName);
                            await vscode.workspace.applyEdit(replaceEdit);
                        }
                    }
                }
            }
        })
    );
    context.subscriptions.push(
        vscode.commands.registerCommand('vba-runner.extractFunction',
            (_uri: string, _procName: string, _startLine: number, _endLine: number) => {
                vscode.window.showInformationMessage('Extract Function は未実装です。');
            }
        )
    );
    outputChannel.appendLine('✓ Code Actions (Refactor) registered');

    outputChannel.appendLine('✓ All providers registered successfully');
    outputChannel.appendLine('📝 Open a .bas file and hover over code to test LSP features');
    console.log('🚀 VBA extension fully initialized');
}

export function deactivate() {
    console.log('VBA Runner extension deactivated');
}

