import * as vscode from 'vscode';
import * as path from 'path';
import * as fs from 'fs';
import { LSPServer } from './lsp/server';
import { VBADebugAdapterFactory } from './lsp/vscode-debug-adapter';
import { Lexer } from './engine/lexer';
import { Parser } from './engine/parser';
import { Evaluator } from './engine/evaluator';

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

    outputChannel.appendLine('✓ All providers registered successfully');
    outputChannel.appendLine('📝 Open a .bas file and hover over code to test LSP features');
    outputChannel.show();
    console.log('🚀 VBA extension fully initialized');
}

export function deactivate() {
    console.log('VBA Runner extension deactivated');
}
