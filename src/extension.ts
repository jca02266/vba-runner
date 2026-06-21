import * as vscode from 'vscode';
import * as l10n from '@vscode/l10n';

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
import { canonicalKeyword, isInStringOrComment } from './lsp/keyword-casing';
import { autoParensEdit, getBlockEnd, needsBodyIndent, needsEndBlock } from './lsp/auto-parens';
import { checkOptionExplicit } from './engine/option-explicit-checker';
import { loadMocks } from '../test-libs/mock-loader';
import { injectExcelStub } from '../test-libs/excel-stub';

let lspServer: LSPServer;
const documentMap = new Map<string, vscode.TextDocument>();
// extension host リロード時に同一 ID での createTestController 二重呼び出しを防ぐ
let _testController: vscode.TestController | undefined;
// activate() の2重呼び出し時に旧プロバイダーを破棄するためモジュールレベルで管理
let _diagnosticCollection: vscode.DiagnosticCollection | undefined;
let _hoverProviderReg: vscode.Disposable | undefined;
let _definitionProviderReg: vscode.Disposable | undefined;
let _referencesProviderReg: vscode.Disposable | undefined;

export async function activate(context: vscode.ExtensionContext) {
    console.log('🚀 VBA Runner extension activated');

    // 前回の activate() で登録した主要プロバイダーを破棄してから再登録する。
    // これにより dev + installed 共存や extension host 不完全リロード時の2重登録を防ぐ。
    _hoverProviderReg?.dispose();
    _definitionProviderReg?.dispose();
    _referencesProviderReg?.dispose();
    _diagnosticCollection?.dispose();

    // Load l10n bundle for the current VS Code language (handles F5 dev mode where vscode.l10n.bundle is undefined)
    const lang = vscode.env.language;
    const bundleFsPath = vscode.Uri.joinPath(context.extensionUri, 'l10n', `bundle.l10n.${lang}.json`).fsPath;
    try {
        l10n.config({ fsPath: bundleFsPath });
    } catch {
        // No bundle for this locale; English strings will be used as-is
    }

    const outputChannel = vscode.window.createOutputChannel('VBA Runner');
    context.subscriptions.push(outputChannel);
    outputChannel.appendLine('Extension initialization started...');

    lspServer = new LSPServer();
    outputChannel.appendLine('LSP Server initialized');

    // ファイル削除・リネーム時にワークスペースキャッシュから除去する
    // （新規ファイルは F12 押下時の遅延スキャンで自動的に取得される）
    context.subscriptions.push(
        vscode.workspace.onDidDeleteFiles((event) => {
            for (const fileUri of event.files) {
                lspServer.unloadWorkspaceFile(fileUri.toString());
            }
        })
    );
    context.subscriptions.push(
        vscode.workspace.onDidRenameFiles((event) => {
            for (const { oldUri } of event.files) {
                lspServer.unloadWorkspaceFile(oldUri.toString());
            }
        })
    );

    // Create diagnostic collection for VBA parse errors
    _diagnosticCollection = vscode.languages.createDiagnosticCollection('vba');
    const diagnosticCollection = _diagnosticCollection;
    context.subscriptions.push(diagnosticCollection);

    // Create diagnostic collection for VBA runtime errors (shown in Problems panel)
    const runtimeDiagnostics = vscode.languages.createDiagnosticCollection('vba-runner');
    context.subscriptions.push(runtimeDiagnostics);

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
            const msg = d.l10nKey != null
                ? l10n.t(d.l10nKey, ...(d.l10nArgs ?? []))
                : d.message;
            const diag = new vscode.Diagnostic(range, msg, sev);
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
    _hoverProviderReg = vscode.languages.registerHoverProvider('vba', {
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
    });
    context.subscriptions.push(_hoverProviderReg);
    outputChannel.appendLine('✓ Hover provider registered');

    // Register definition provider
    _definitionProviderReg = vscode.languages.registerDefinitionProvider('vba', {
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
    });
    context.subscriptions.push(_definitionProviderReg);
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
    _referencesProviderReg = vscode.languages.registerReferenceProvider('vba', {
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
    });
    context.subscriptions.push(_referencesProviderReg);
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
                    if (item.insertText !== undefined) ci.insertText = item.insertText;
                    if (item.filterText !== undefined) ci.filterText = item.filterText;
                    if (item.sortText !== undefined) ci.sortText = item.sortText;
                    if (item.replaceStartCharacter !== undefined) {
                        ci.range = new vscode.Range(
                            position.line, item.replaceStartCharacter,
                            position.line, position.character
                        );
                    }
                    return ci;
                });
            }
        }, '.', ' ')); // Trigger on dot and space (for "End ", "Next ", etc.)
    outputChannel.appendLine('✓ Completion provider registered');

    // Register test discovery
    // モジュールレベルで保持し、リロード時に旧インスタンスを破棄してから再生成する
    if (_testController) {
        _testController.dispose();
    }
    _testController = vscode.tests.createTestController('vbaRunner', 'VBA Tests');
    const testController = _testController;
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
            // catch ブロックからも参照できるようスコープを外に出す
            const moduleFileMap = new Map<string, string>();
            try {
                const doc = documentMap.get(uri);
                if (!doc) {
                    outputChannel.appendLine(`[Error] Document not found in map: ${uri}`);
                    outputChannel.show();
                    return;
                }

                // 同ディレクトリの .bas/.cls をマルチモジュール評価（行番号をファイル単位で独立させる）
                const dir = path.dirname(doc.uri.fsPath);
                const entries = fs.readdirSync(dir).filter(f => /\.(bas|cls)$/i.test(f));

                const ev = new Evaluator((msg: string) => outputChannel.appendLine(msg), { allowTopLevelStatements: false });
                // Excel API スタブを自動注入（Range / Cells / ActiveSheet 等をノーオプで動かす）
                injectExcelStub(ev);
                const asts: Array<{ ast: ReturnType<Parser['parse']>; moduleName: string }> = [];

                // AssertHelper をクラスモジュールとして評価
                const assertAst = new Parser(
                    new Lexer(ASSERT_HELPER_SRC).tokenize(),
                    { parseAsClass: 'AssertHelper' }
                ).parse();
                ev.evaluateModule(assertAst);
                asts.push({ ast: assertAst, moduleName: 'AssertHelper' });

                // __mocks__/ と __mocks__.* からモックを注入（本番モジュールより先）
                const mockModules = loadMocks(dir, ev);
                for (const { ast: mockAst, moduleName: mockName } of mockModules) {
                    asts.push({ ast: mockAst, moduleName: mockName });
                }

                for (const entry of entries) {
                    // ディレクトリに AssertHelper.cls があっても重複注入しない
                    if (/^AssertHelper\.cls$/i.test(entry)) continue;
                    const moduleName = path.basename(entry, path.extname(entry));
                    const filePath = path.join(dir, entry);
                    moduleFileMap.set(moduleName.toLowerCase(), filePath);
                    // エディター上で開いている場合は未保存変更を含むメモリ上のテキストを優先する
                    const fileUriStr = vscode.Uri.file(filePath).toString();
                    const openDoc = documentMap.get(fileUriStr);
                    const content = openDoc ? openDoc.getText() : fs.readFileSync(filePath, 'utf8');
                    const tokens = new Lexer(content).tokenize();
                    let ast;
                    if (/\.cls$/i.test(entry)) {
                        ast = new Parser(tokens, { parseAsClass: moduleName }).parse();
                        ev.evaluateModule(ast);
                    } else {
                        ev.setSourceModule(moduleName);
                        ast = new Parser(tokens, { errorRecovery: true }).parse();
                        for (const d of ast.diagnostics) {
                            outputChannel.appendLine(`[parse warning] ${moduleName} line ${d.loc.start.line}: ${d.message}`);
                        }
                        ev.evaluateModule(ast);
                    }
                    asts.push({ ast, moduleName });
                }

                let callTarget = procName;
                if (isTestProc) {
                    // ラッパー Sub を生成: assert インスタンスを作って渡し、結果を表示
                    const wrapperName = `__VBARunner_${procName}__`;
                    const wrapperSrc = [
                        `Sub ${wrapperName}()`,
                        `    Dim assert As New AssertHelper`,
                        `    ${procName} assert`,
                        `    If assert.Failed Then`,
                        `        Debug.Print "[FAIL] ${procName}: " & assert.FailMessage`,
                        `    Else`,
                        `        Debug.Print "[PASS] ${procName}"`,
                        `    End If`,
                        `End Sub`,
                    ].join('\n');
                    ev.setSourceModule('__VBARunner_wrapper__');
                    const wrapperAst = new Parser(new Lexer(wrapperSrc).tokenize()).parse();
                    ev.evaluateModule(wrapperAst);
                    asts.push({ ast: wrapperAst, moduleName: '__VBARunner_wrapper__' });
                    callTarget = wrapperName;
                }

                ev.reEvaluateModuleConstsAll(asts);
                runtimeDiagnostics.clear();
                const result = ev.callProcedure(callTarget, []);
                if (result !== undefined) {
                    outputChannel.appendLine(`[Run] ${procName}() → ${result}`);
                }
                outputChannel.show();
            } catch (e: any) {
                outputChannel.appendLine(`[Error] ${procName}: ${e.message}`);

                // スタックトレースを Output にも出力する。
                // frames[i].line = frame[i] が呼び出し元（frames[i+1]）の何行目から呼ばれたか。
                // よって frame[i] 内で実行されていた行は:
                //   - 最深フレーム frame[0]: エラー発生行 e.vbaLine
                //   - それ以外 frame[i]: frames[i-1].line（次の深いフレームを呼び出した行）
                const frames: Array<{ name: string; moduleName: string; line: number }> =
                    Array.isArray(e.vbaStack) ? e.vbaStack : [];
                const frameLine = (i: number): number =>
                    i === 0 ? (e.vbaLine || 0) : frames[i - 1].line;
                for (let i = 0; i < frames.length; i++) {
                    const mod = frames[i].moduleName ? `${frames[i].moduleName}.` : '';
                    const ln = frameLine(i);
                    outputChannel.appendLine(`    at ${mod}${frames[i].name}${ln ? ` (line ${ln})` : ''}`);
                }

                if (e.vbaLine && e.vbaModule) {
                    const filePath = moduleFileMap.get((e.vbaModule as string).toLowerCase());
                    if (filePath) {
                        const line = (e.vbaLine as number) - 1;
                        const uri = vscode.Uri.file(filePath);
                        const diag = new vscode.Diagnostic(
                            new vscode.Range(line, 0, line, 999),
                            `${procName}: ${e.message}`,
                            vscode.DiagnosticSeverity.Error
                        );

                        // スタックトレースを relatedInformation（Problems パネル上の
                        // クリック可能なサブ項目）として添付する。各フレームの該当行へジャンプできる。
                        // frames[0]（最内フレーム）はメイン diagnostic と同じ行を指すため除外する。
                        const related: vscode.DiagnosticRelatedInformation[] = [];
                        for (let i = 1; i < frames.length; i++) {
                            const fp = moduleFileMap.get(frames[i].moduleName.toLowerCase());
                            if (!fp) continue;
                            const fl = Math.max(0, frameLine(i) - 1);
                            const mod = frames[i].moduleName ? `${frames[i].moduleName}.` : '';
                            related.push(new vscode.DiagnosticRelatedInformation(
                                new vscode.Location(vscode.Uri.file(fp), new vscode.Range(fl, 0, fl, 999)),
                                `at ${mod}${frames[i].name}`
                            ));
                        }
                        if (related.length > 0) diag.relatedInformation = related;

                        runtimeDiagnostics.set(uri, [diag]);
                        // Problems にフォーカス（outputChannel.show() は呼ばない）
                        vscode.commands.executeCommand('workbench.panel.markers.view.focus');
                        return;
                    }
                }
                // ファイル情報がない場合は Output にフォールバック
                outputChannel.show();
            }
        })
    );

    // vba-runner.debugProcedure: CodeLens から呼ばれる「指定プロシージャをデバッグ実行」コマンド
    context.subscriptions.push(
        vscode.commands.registerCommand('vba-runner.debugProcedure', (uri: string, procName: string) => {
            const filePath = vscode.Uri.parse(uri).fsPath;
            const folder = vscode.workspace.getWorkspaceFolder(vscode.Uri.parse(uri));
            vscode.debug.startDebugging(folder, {
                type: 'vba',
                request: 'launch',
                name: `Debug VBA: ${procName}`,
                program: filePath,
                entryPoint: procName,
            });
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

    // vba-runner.generateTest: Code Lens「未テスト」から呼ばれる
    context.subscriptions.push(
        vscode.commands.registerCommand('vba-runner.generateTest', async (uri: string, procName: string) => {
            if (!uri || !procName) {
                vscode.window.showErrorMessage(l10n.t('Open a VBA file and run from the procedure Code Lens'));
                return;
            }

            const docUri = vscode.Uri.parse(uri);
            const sourceFilePath = docUri.fsPath;
            const sourceDir = path.dirname(sourceFilePath);
            const sourceBase = path.basename(sourceFilePath, path.extname(sourceFilePath));

            // --- テスト配置設定の読み込み or 初回選択 ---
            type TestLocation = 'sameFile' | 'separateFile';
            const config = vscode.workspace.getConfiguration('vba-runner');
            let testLocation = config.get<TestLocation>('test.location');
            if (!testLocation) {
                const choice = await vscode.window.showQuickPick(
                    [
                        { label: l10n.t('Add to same file'), description: l10n.t('Append Test_MethodName at the end of the same .bas file'), value: 'sameFile' as TestLocation },
                        { label: l10n.t('Add to separate file'), description: l10n.t('Create {0}Test.bas and append the test', sourceBase), value: 'separateFile' as TestLocation },
                    ],
                    { placeHolder: l10n.t('Select test location (saved to workspace settings)') }
                );
                if (!choice) return;
                testLocation = choice.value;
                await config.update('test.location', testLocation, vscode.ConfigurationTarget.Workspace);
            }

            // --- テストスタブ文字列の生成 ---
            const buildStub = (targetUri: string): string => {
                const src = fs.readFileSync(vscode.Uri.parse(targetUri).fsPath, 'utf-8');
                const ast = lspServer.parseDocument(src);
                let paramNames: string[] = [];
                let isFunction = false;
                if (ast?.body) {
                    const procDecl = (ast.body as any[]).find(
                        (s: any) => s.type === 'ProcedureDeclaration' &&
                            s.name?.name?.toLowerCase() === procName.toLowerCase()
                    );
                    if (procDecl) {
                        paramNames = (procDecl.parameters ?? []).map((p: any) => String(p.name));
                        isFunction = procDecl.isFunction ?? false;
                    }
                }
                const callArgs = paramNames.join(', ');
                const callExpr = isFunction
                    ? `result = ${procName}(${callArgs})`
                    : `${procName}${callArgs ? ' ' + callArgs : ''}`;
                const lines = ['', `Sub Test_${procName}(assert)`, `    ' TODO: ${l10n.t('Implement the test')}`];
                if (isFunction) {
                    lines.push(`    ' Dim result`, `    ' ${callExpr}`, `    ' assert.IsTrue result = expected, "${l10n.t('description')}"`);
                } else {
                    lines.push(`    ' ${callExpr}`, `    ' assert.IsTrue condition, "${l10n.t('description')}"`);
                }
                lines.push('End Sub');
                return lines.join('\n');
            };

            // テスト済みかチェックしてカーソル移動、未テストならスタブを追記
            const navigateOrInsert = async (targetDocUri: vscode.Uri) => {
                const testName = `Test_${procName}`;
                const targetUriStr = targetDocUri.toString();
                const symbols = lspServer.getDocumentSymbols(targetUriStr);
                const existing = symbols.find((s: any) => s.name?.toLowerCase() === testName.toLowerCase());
                if (existing) {
                    const pos = new vscode.Position(existing.location.range.start.line, 0);
                    await vscode.window.showTextDocument(targetDocUri, { selection: new vscode.Range(pos, pos) });
                    return;
                }
                const doc = await vscode.workspace.openTextDocument(targetDocUri);
                const lastLine = doc.lineCount - 1;
                const lastChar = doc.lineAt(lastLine).text.length;
                const edit = new vscode.WorkspaceEdit();
                edit.insert(targetDocUri, new vscode.Position(lastLine, lastChar), buildStub(uri));
                await vscode.workspace.applyEdit(edit);
                await vscode.workspace.openTextDocument(targetDocUri);
                const newSymbols = lspServer.getDocumentSymbols(targetUriStr);
                const newSym = newSymbols.find((s: any) => s.name?.toLowerCase() === testName.toLowerCase());
                const targetLine = newSym ? newSym.location.range.start.line : lastLine + 2;
                const pos = new vscode.Position(targetLine, 0);
                await vscode.window.showTextDocument(targetDocUri, { selection: new vscode.Range(pos, pos) });
                vscode.window.showInformationMessage(l10n.t("Generated stub for '{0}'", testName));
            };

            if (testLocation === 'sameFile') {
                await navigateOrInsert(docUri);
            } else {
                const testFilePath = path.join(sourceDir, `${sourceBase}Test.bas`);
                const testFileUri = vscode.Uri.file(testFilePath);
                if (!fs.existsSync(testFilePath)) {
                    fs.writeFileSync(testFilePath, 'Option Explicit\n', 'utf-8');
                    // LSP に新ファイルを認識させる
                    const newDoc = await vscode.workspace.openTextDocument(testFileUri);
                    lspServer.didOpen(testFileUri.toString(), newDoc.getText());
                }
                await navigateOrInsert(testFileUri);
            }
        })
    );

    // vba-runner.generateMocks: __mocks__/ExcelObjects.bas のひな形を生成
    context.subscriptions.push(
        vscode.commands.registerCommand('vba-runner.generateMocks', async () => {
            const editor = vscode.window.activeTextEditor;
            if (!editor) {
                vscode.window.showErrorMessage(l10n.t('Open a VBA file before running this command'));
                return;
            }

            const filePath = editor.document.uri.fsPath;
            const dir = path.dirname(filePath);
            const outPath = path.join(dir, '__mocks__', 'ExcelObjects.bas');

            try {
                // vba-analyzer の解析ロジックをプロセス内で直接呼び出す
                // （execSync + npx tsx での子プロセス起動は、配布後の .vsix に test-libs/ が
                //   同梱されないため動作しない。esbuild が静的 import を extension.cjs に
                //   バンドルするため、この形なら配布後も動作する）
                const { collectVbaFilesForMcp, analyzeWorkspaceForMcpFromFiles } = await import('../test-libs/vba-analyzer');
                const files = collectVbaFilesForMcp(dir);
                const json = analyzeWorkspaceForMcpFromFiles(files);
                const { extractObjectsFromAnalyzerJson, generateExcelMockBas } = await import('../test-libs/mock-generator');
                const objects = extractObjectsFromAnalyzerJson(json);

                if (objects.length === 0) {
                    vscode.window.showInformationMessage(l10n.t('No Excel-dependent objects detected.'));
                    return;
                }

                // 既存ファイルのチェック
                if (fs.existsSync(outPath)) {
                    const overwrite = l10n.t('Overwrite');
                    const answer = await vscode.window.showWarningMessage(
                        l10n.t('__mocks__/ExcelObjects.bas already exists. Overwrite?'),
                        overwrite, l10n.t('Cancel')
                    );
                    if (answer !== overwrite) return;
                }

                // ひな形生成・書き出し
                const content = generateExcelMockBas(objects, {
                    procName: path.basename(filePath),
                    date: new Date().toISOString().slice(0, 10),
                });
                fs.mkdirSync(path.dirname(outPath), { recursive: true });
                fs.writeFileSync(outPath, content, 'utf-8');

                // 生成したファイルを開く
                const doc = await vscode.workspace.openTextDocument(outPath);
                await vscode.window.showTextDocument(doc);
                vscode.window.showInformationMessage(
                    l10n.t('Detected: {0}\nGenerated __mocks__/ExcelObjects.bas.', objects.join(', '))
                );
            } catch (e: any) {
                vscode.window.showErrorMessage(l10n.t('generateMocks error: {0}', e.message));
                outputChannel.appendLine(`[generateMocks] ${e.message}`);
            }
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
                vscode.window.showInformationMessage(l10n.t("Test function '{0}' not found", testName));
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

    // Register on-type formatting provider:
    //   - auto line continuation: append ' _' when an expression is left dangling
    //   - auto keyword casing: rewrite keywords to canonical VBA casing (VBE-style)
    context.subscriptions.push(
        vscode.languages.registerOnTypeFormattingEditProvider('vba', {
            provideOnTypeFormattingEdits(document, position, ch) {
                const config = vscode.workspace.getConfiguration('vba-runner');
                const edits: vscode.TextEdit[] = [];

                if (ch === '\n' && position.line > 0) {
                    const prevLine = document.lineAt(position.line - 1);
                    const trimmed = prevLine.text.trimEnd();

                    // Auto line continuation: insert ' _' at the end of the previous line.
                    if (config.get('editor.autoLineContinuation', true)) {
                        if (needsLineContinuation(trimmed)) {
                            const insertPos = new vscode.Position(position.line - 1, trimmed.length);
                            const lineEnd = new vscode.Position(position.line - 1, prevLine.text.length);
                            edits.push(vscode.TextEdit.replace(new vscode.Range(insertPos, lineEnd), ' _'));
                        }
                    }

                    // Auto parentheses and auto end-block are computed together so they can
                    // share a single editor.edit call when both apply. Mixing TextEdit (from
                    // the provider return value) with a concurrent editor.edit (setTimeout)
                    // causes VS Code to reject the deferred edit, making the two features
                    // mutually exclusive. One editor.edit handles both atomically.
                    const ap = config.get('editor.autoParentheses', true)
                        ? autoParensEdit(prevLine.text) : null;

                    const autoEndBlock = config.get('editor.autoEndBlock', true);
                    const blockEnd = autoEndBlock ? getBlockEnd(prevLine.text) : null;
                    let shouldInsertEnd = false;
                    if (blockEnd) {
                        const getLine = (n: number) =>
                            n < document.lineCount ? document.lineAt(n).text : undefined;
                        shouldInsertEnd = needsEndBlock(
                            getLine, position.line + 1,
                            blockEnd.closePattern, blockEnd.openPattern);
                    }
                    // Else / ElseIf / Case: indent body without inserting a new end keyword.
                    const indentOnly = autoEndBlock && !shouldInsertEnd
                        && needsBodyIndent(prevLine.text);

                    if (ap && !shouldInsertEnd && !indentOnly) {
                        // Only parens needed: fast path via TextEdit (single atomic edit).
                        edits.push(vscode.TextEdit.insert(
                            new vscode.Position(position.line - 1, ap.insertCol), '()'));
                    } else if (shouldInsertEnd || indentOnly) {
                        // End block / indent-only (+ possibly parens): combine into one
                        // editor.edit to avoid conflicts with concurrent TextEdits.
                        const insertKeyword = shouldInsertEnd ? blockEnd!.insertKeyword : null;
                        const parenCol = (ap && shouldInsertEnd) ? ap.insertCol : -1;
                        const prevLineIdx = position.line - 1;
                        const baseIndent = prevLine.text.match(/^(\s*)/)?.[1] ?? '';
                        const docUri = document.uri.toString();
                        setTimeout(() => {
                            const editor = vscode.window.activeTextEditor;
                            if (!editor) return;
                            if (editor.document.uri.toString() !== docUri) return;
                            const cursorLine = editor.selection.active.line;
                            const lineRange = editor.document.lineAt(cursorLine).range;
                            const tabSize = typeof editor.options.tabSize === 'number'
                                ? editor.options.tabSize : 4;
                            const indentUnit = editor.options.insertSpaces !== false
                                ? ' '.repeat(tabSize) : '\t';
                            const bodyIndent = baseIndent + indentUnit;
                            editor.edit(eb => {
                                if (parenCol >= 0) {
                                    eb.insert(new vscode.Position(prevLineIdx, parenCol), '()');
                                }
                                if (insertKeyword) {
                                    eb.replace(lineRange, bodyIndent + '\n' + baseIndent + insertKeyword);
                                } else {
                                    eb.replace(lineRange, bodyIndent);
                                }
                            }).then(() => {
                                const newPos = new vscode.Position(cursorLine, bodyIndent.length);
                                editor.selection = new vscode.Selection(newPos, newPos);
                            });
                        }, 0);
                    }
                }

                // Auto keyword casing: rewrite the just-completed word to its
                // canonical VBA casing (e.g. 'if' -> 'If', 'withevents' -> 'WithEvents').
                if (config.get('editor.autoKeywordCasing', true)) {
                    const casing = keywordCasingEdit(document, position, ch);
                    // Avoid overlapping the line-continuation edit (different ranges in practice).
                    if (casing && !edits.some((e) => !!e.range.intersection(casing.range))) {
                        edits.push(casing);
                    }
                }

                return edits;
            }
        }, '\n', ' ', '\t')
    );
    outputChannel.appendLine('✓ On-type formatting provider (line continuation + keyword casing) registered');

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
                    const introduceVarAction = new vscode.CodeAction('Introduce Variable', vscode.CodeActionKind.Refactor);
                    introduceVarAction.command = {
                        title: 'Introduce Variable',
                        command: 'vba-runner.introduceVariable',
                        arguments: [document.uri, range],
                    };
                    actions.push(introduceVarAction);

                    // Extract Constant: selection is a literal value
                    const selectedText = document.getText(range).trim();
                    if (/^("(?:[^""]|"")*"|-?\d+(?:\.\d+)?(?:[Ee][+-]?\d+)?|True|False)$/i.test(selectedText)) {
                        const extractConstAction = new vscode.CodeAction('Extract Constant', vscode.CodeActionKind.RefactorExtract);
                        extractConstAction.command = {
                            title: 'Extract Constant',
                            command: 'vba-runner.extractConstant',
                            arguments: [document.uri, range],
                        };
                        actions.push(extractConstAction);
                    }

                    // Introduce With: selection spans multiple lines with same object prefix
                    if (range.start.line < range.end.line) {
                        const lineTexts: string[] = [];
                        for (let i = range.start.line; i <= range.end.line; i++) {
                            lineTexts.push(document.lineAt(i).text.trim());
                        }
                        const nonEmpty = lineTexts.filter(Boolean);
                        const objMatch = nonEmpty[0]?.match(/^([A-Za-z_][A-Za-z0-9_]*(?:\.[A-Za-z_][A-Za-z0-9_]*)*)\./);
                        if (objMatch && nonEmpty.every(l => l.toLowerCase().startsWith(objMatch[1].toLowerCase() + '.'))) {
                            const withAction = new vscode.CodeAction(
                                `Introduce With '${objMatch[1]}'`,
                                vscode.CodeActionKind.Refactor
                            );
                            withAction.command = {
                                title: 'Introduce With',
                                command: 'vba-runner.introduceWith',
                                arguments: [document.uri, range, objMatch[1]],
                            };
                            actions.push(withAction);
                        }
                    }

                    // Extract Sub/Function (from LSP server)
                    const lspActions = lspServer.getCodeActions(document.uri.toString(), {
                        start: { line: range.start.line, character: range.start.character },
                        end:   { line: range.end.line,   character: range.end.character },
                    });
                    for (const la of lspActions) {
                        const vsAction = new vscode.CodeAction(la.title, vscode.CodeActionKind.RefactorExtract);
                        vsAction.command = la.command;
                        actions.push(vsAction);
                    }
                } else {
                    // Inline Variable: cursor on an identifier (no selection)
                    const wordRange = document.getWordRangeAtPosition(range.start, /[A-Za-z_][A-Za-z0-9_]*/);
                    if (wordRange) {
                        const word = document.getText(wordRange);
                        if (!/^(Sub|Function|End|Dim|As|Private|Public|Friend|If|Then|Else|ElseIf|For|Each|In|Next|Do|While|Until|Loop|With|Select|Case|GoTo|GoSub|Return|Exit|True|False|Nothing|Null|Empty|Not|And|Or|Xor|Eqv|Imp|Mod|Like|Is|Let|Set|New|ByVal|ByRef|Optional|ParamArray|Preserve|Static|Type|Enum|Property|Get|Put|Call|On|Error|Resume|Wend|Me|ReDim)$/i.test(word)) {
                            const inlineAction = new vscode.CodeAction(
                                `Inline Variable '${word}'`,
                                vscode.CodeActionKind.RefactorInline
                            );
                            inlineAction.command = {
                                title: 'Inline Variable',
                                command: 'vba-runner.inlineVariable',
                                arguments: [document.uri, wordRange],
                            };
                            actions.push(inlineAction);
                        }
                    }
                }

                return actions;
            }
        }, {
            providedCodeActionKinds: [
                vscode.CodeActionKind.Refactor,
                vscode.CodeActionKind.RefactorExtract,
                vscode.CodeActionKind.RefactorInline,
            ]
        })
    );

    // Register Source Actions (Source Actions menu)
    context.subscriptions.push(
        vscode.languages.registerCodeActionsProvider('vba', {
            provideCodeActions(document) {
                const actions: vscode.CodeAction[] = [];

                const removeAction = new vscode.CodeAction(
                    'Remove Unused Variables',
                    vscode.CodeActionKind.SourceFixAll
                );
                removeAction.command = {
                    title: 'Remove Unused Variables',
                    command: 'vba-runner.removeUnusedVariables',
                    arguments: [document.uri],
                };
                actions.push(removeAction);

                const organizeAction = new vscode.CodeAction(
                    'Organize Declarations',
                    vscode.CodeActionKind.Source
                );
                organizeAction.command = {
                    title: 'Organize Declarations',
                    command: 'vba-runner.organizeDeclarations',
                    arguments: [document.uri],
                };
                actions.push(organizeAction);

                return actions;
            }
        }, {
            providedCodeActionKinds: [
                vscode.CodeActionKind.Source,
                vscode.CodeActionKind.SourceFixAll,
            ]
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
                const REPLACE_ALL = l10n.t('Replace all');
                const REVIEW_ONE = l10n.t('Review one by one');
                const choice = await vscode.window.showQuickPick(
                    [REPLACE_ALL, REVIEW_ONE, l10n.t('Keep as is')],
                    { placeHolder: l10n.t('{0} more occurrence(s) found', String(otherReplacements.length)) }
                );

                if (choice === REPLACE_ALL) {
                    // Apply all replacements from back to front to avoid offset changes
                    otherReplacements.sort((a, b) => b.line - a.line || b.start - a.start);
                    for (const repl of otherReplacements) {
                        const replaceEdit = new vscode.WorkspaceEdit();
                        replaceEdit.replace(uri, new vscode.Range(repl.line, repl.start, repl.line, repl.end), varName);
                        await vscode.workspace.applyEdit(replaceEdit);
                    }
                } else if (choice === REVIEW_ONE) {
                    // Apply replacements one by one, asking for each
                    otherReplacements.sort((a, b) => b.line - a.line || b.start - a.start);
                    for (const repl of otherReplacements) {
                        const lineText = editor.document.lineAt(repl.line).text;
                        const exprText = lineText.substring(repl.start, repl.end);

                        const REPLACE = l10n.t('Replace');
                        const userChoice = await vscode.window.showQuickPick(
                            [REPLACE, l10n.t('Skip')],
                            {
                                placeHolder: l10n.t('Line {0}: Replace "{1}"?', String(repl.line + 1), exprText),
                                canPickMany: false
                            }
                        );

                        if (userChoice === REPLACE) {
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
        vscode.commands.registerCommand('vba-runner.doExtractFunction',
            async (
                uri: string,
                lspRange: { start: { line: number; character: number }; end: { line: number; character: number } },
                result: { inputs: string[]; outputs: string[]; locals: string[] },
                procSignature: string,
                callStatement: string,
            ) => {
                const vsUri = vscode.Uri.parse(uri);
                const editor = vscode.window.activeTextEditor;
                if (!editor || editor.document.uri.toString() !== vsUri.toString()) {
                    vscode.window.showWarningMessage(l10n.t('Document context has been lost'));
                    return;
                }

                const defaultName = callStatement.match(/^([a-zA-Z_][a-zA-Z0-9_]*)/)?.[1] ?? 'ExtractedSub';
                const procName = await vscode.window.showInputBox({
                    prompt: l10n.t('New procedure name:'),
                    value: defaultName,
                    validateInput: (input) => {
                        if (!input) return l10n.t('Procedure name is required');
                        if (!/^[a-zA-Z_][a-zA-Z0-9_]*$/.test(input)) return l10n.t('Invalid procedure name');
                        return '';
                    },
                });
                if (!procName) return;

                const startLine = lspRange.start.line;
                const endLine   = lspRange.end.line;

                // Collect selected lines
                const selectedLines: string[] = [];
                for (let i = startLine; i <= endLine; i++) {
                    selectedLines.push(editor.document.lineAt(i).text);
                }

                // Re-indent: strip common leading whitespace, add 4-space indent
                const nonBlank = selectedLines.filter(l => l.trim().length > 0);
                const minIndentLen = nonBlank.length > 0
                    ? Math.min(...nonBlank.map(l => (l.match(/^(\s*)/)?.[1].length) ?? 0))
                    : 0;
                const reindented = selectedLines.map(l =>
                    l.trim().length > 0 ? '    ' + l.slice(minIndentLen) : ''
                );

                // Apply user-provided name to signature/call
                const finalSigLine = procSignature.split('\n')[0].replace(/\bExtractedSub\b/, procName);
                const finalCall    = callStatement.replace(/^ExtractedSub\b/, procName);

                // Add Dim declarations for locals not already declared within selected lines
                const dimmedInSelection = new Set(
                    selectedLines
                        .map(l => l.match(/^\s*Dim\s+(\w+)/i)?.[1]?.toLowerCase())
                        .filter((v): v is string => v !== undefined)
                );
                const extraDims = result.locals
                    .filter(v => !dimmedInSelection.has(v.toLowerCase()))
                    .map(v => `    Dim ${v} As Variant`);

                // Build new Sub text
                const newProcText = [finalSigLine, ...extraDims, ...reindented, 'End Sub'].join('\n');

                // Find containing procedure's last line (0-based) for insertion point
                const ast = lspServer.parseDocument(editor.document.getText());
                let procEndLine = endLine;
                if (ast?.body) {
                    for (const stmt of ast.body) {
                        if (
                            stmt.type === 'ProcedureDeclaration' &&
                            stmt.loc != null &&
                            stmt.loc.start.line - 1 <= startLine &&
                            stmt.loc.end.line   - 1 >= endLine
                        ) {
                            procEndLine = stmt.loc.end.line - 1;
                            break;
                        }
                    }
                }

                const callIndent = selectedLines[0]?.match(/^\s*/)?.[0] ?? '';

                const edit = new vscode.WorkspaceEdit();
                edit.replace(
                    vsUri,
                    new vscode.Range(startLine, 0, endLine, editor.document.lineAt(endLine).text.length),
                    callIndent + finalCall,
                );
                edit.insert(vsUri, new vscode.Position(procEndLine + 1, 0), '\n' + newProcText + '\n');
                await vscode.workspace.applyEdit(edit);
            }
        )
    );
    context.subscriptions.push(
        vscode.commands.registerCommand('vba-runner.extractFunction',
            (_uri: string, _procName: string, _startLine: number, _endLine: number) => {
                vscode.window.showInformationMessage(l10n.t('Extract Function is not yet implemented.'));
            }
        )
    );

    // Extract Constant: selected literal → Const at procedure top, replace all occurrences
    context.subscriptions.push(
        vscode.commands.registerCommand('vba-runner.extractConstant', async (uri: vscode.Uri, range: vscode.Range) => {
            const editor = vscode.window.activeTextEditor;
            if (!editor || editor.document.uri.toString() !== uri.toString()) return;

            const literal = editor.document.getText(range).trim();

            // Derive default name from literal
            let defaultName = 'CONST_VALUE';
            if (literal.startsWith('"')) {
                const word = literal.slice(1, -1).trim().match(/^[A-Za-z_][A-Za-z0-9_]*/)?.[0];
                if (word) defaultName = word.toUpperCase();
            } else if (/^\d/.test(literal)) {
                defaultName = 'CONST_' + literal.replace(/[^0-9A-Za-z]/g, '_').replace(/_+$/, '');
            }

            const constName = await vscode.window.showInputBox({
                prompt: l10n.t('Constant name:'),
                value: defaultName,
                validateInput: v => !v ? l10n.t('Constant name is required') : !/^[A-Za-z_][A-Za-z0-9_]*$/.test(v) ? l10n.t('Invalid constant name') : '',
            });
            if (!constName) return;

            const ast = lspServer.parseDocument(editor.document.getText());
            const lineNum1 = range.start.line + 1; // 1-based
            const proc = ast?.body?.find((s: any) =>
                s.type === 'ProcedureDeclaration' && s.loc?.start.line <= lineNum1 && s.loc?.end.line >= lineNum1
            );

            if (!proc) {
                vscode.window.showWarningMessage(l10n.t('The cursor must be inside a procedure'));
                return;
            }

            const procStart0 = (proc.loc.start.line as number) - 1; // 0-based
            const procEnd0 = (proc.loc.end.line as number) - 1;
            const insertLine0 = procStart0 + 1; // line after Sub/Function header
            const insertIndent = editor.document.lineAt(insertLine0).text.match(/^\s*/)?.[0] ?? '    ';

            // Build search pattern for this literal
            const isString = literal.startsWith('"');
            const escaped = literal.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
            const pattern = isString
                ? new RegExp(escaped, 'g')
                : new RegExp(`(?<![A-Za-z0-9_.])${escaped}(?![A-Za-z0-9_.])`, 'gi');

            // Collect all occurrences in the procedure body (skip comments)
            const replacements: vscode.Range[] = [];
            for (let i = procStart0 + 1; i <= procEnd0; i++) {
                const lineText = editor.document.lineAt(i).text;
                const commentIdx = lineText.indexOf("'");
                const searchIn = commentIdx >= 0 ? lineText.slice(0, commentIdx) : lineText;
                pattern.lastIndex = 0;
                let m: RegExpExecArray | null;
                while ((m = pattern.exec(searchIn)) !== null) {
                    replacements.push(new vscode.Range(i, m.index, i, m.index + m[0].length));
                }
            }

            const edit = new vscode.WorkspaceEdit();
            // Apply replacements bottom to top (WorkspaceEdit handles this correctly, but explicit ordering is safe)
            for (const r of [...replacements].reverse()) {
                edit.replace(uri, r, constName);
            }
            edit.insert(uri, new vscode.Position(insertLine0, 0), `${insertIndent}Const ${constName} = ${literal}\n`);
            await vscode.workspace.applyEdit(edit);
        })
    );

    // Inline Variable: replace all references with the assigned expression, remove Dim + assignment
    context.subscriptions.push(
        vscode.commands.registerCommand('vba-runner.inlineVariable', async (uri: vscode.Uri, wordRange: vscode.Range) => {
            const editor = vscode.window.activeTextEditor;
            if (!editor || editor.document.uri.toString() !== uri.toString()) return;

            const varName = editor.document.getText(wordRange);
            const ast = lspServer.parseDocument(editor.document.getText());
            const lineNum1 = wordRange.start.line + 1; // 1-based

            const proc = ast?.body?.find((s: any) =>
                s.type === 'ProcedureDeclaration' && s.loc?.start.line <= lineNum1 && s.loc?.end.line >= lineNum1
            );
            if (!proc) {
                vscode.window.showWarningMessage(l10n.t('Variable is not inside a procedure'));
                return;
            }

            const procStart0 = (proc.loc.start.line as number) - 1;
            const procEnd0 = (proc.loc.end.line as number) - 1;

            const dimRe = new RegExp(`^\\s*(?:Static\\s+)?Dim\\s+${varName}\\b`, 'i');
            const assignRe = new RegExp(`^(\\s*)${varName}\\s*=\\s*(.+)$`, 'i');

            let dimLine = -1;
            let assignLine = -1;
            let assignExpr = '';
            let assignCount = 0;

            for (let i = procStart0; i <= procEnd0; i++) {
                const text = editor.document.lineAt(i).text;
                if (dimRe.test(text)) { dimLine = i; continue; }
                const m = text.match(assignRe);
                if (m) {
                    assignCount++;
                    if (assignCount === 1) { assignLine = i; assignExpr = m[2].trim(); }
                }
            }

            if (dimLine < 0) {
                vscode.window.showWarningMessage(l10n.t("No Dim declaration found for '{0}'", varName));
                return;
            }
            if (assignLine < 0) {
                vscode.window.showWarningMessage(l10n.t("No assignment found for '{0}'", varName));
                return;
            }
            if (assignCount > 1) {
                vscode.window.showWarningMessage(l10n.t("Cannot inline '{0}': multiple assignments found", varName));
                return;
            }

            // Collect reference positions (exclude Dim and assignment lines)
            const refPattern = new RegExp(`(?<![A-Za-z0-9_.])${varName}(?![A-Za-z0-9_])`, 'gi');
            const replacements: vscode.Range[] = [];
            for (let i = procStart0; i <= procEnd0; i++) {
                if (i === dimLine || i === assignLine) continue;
                const text = editor.document.lineAt(i).text;
                refPattern.lastIndex = 0;
                let m: RegExpExecArray | null;
                while ((m = refPattern.exec(text)) !== null) {
                    replacements.push(new vscode.Range(i, m.index, i, m.index + m[0].length));
                }
            }

            // Parenthesize if multi-use and expression contains operators
            const needsParens =
                replacements.length > 1 &&
                /[+\-*&^\\]|\bMod\b|\bAnd\b|\bOr\b|\bXor\b|\bNot\b/i.test(assignExpr);
            const inlined = needsParens ? `(${assignExpr})` : assignExpr;

            const edit = new vscode.WorkspaceEdit();
            for (const r of [...replacements].reverse()) {
                edit.replace(uri, r, inlined);
            }
            // Delete lines from bottom to top
            for (const ln of [dimLine, assignLine].sort((a, b) => b - a)) {
                edit.delete(uri, new vscode.Range(ln, 0, ln + 1, 0));
            }
            await vscode.workspace.applyEdit(edit);
        })
    );

    // Introduce With: wrap selected consecutive obj.Xxx lines in With obj ... End With
    context.subscriptions.push(
        vscode.commands.registerCommand('vba-runner.introduceWith', async (uri: vscode.Uri, range: vscode.Range, obj: string) => {
            const editor = vscode.window.activeTextEditor;
            if (!editor || editor.document.uri.toString() !== uri.toString()) return;

            const indent = editor.document.lineAt(range.start.line).text.match(/^\s*/)?.[0] ?? '';
            const innerIndent = indent + '    ';
            const objLower = obj.toLowerCase();

            const bodyLines: string[] = [];
            for (let i = range.start.line; i <= range.end.line; i++) {
                const trimmed = editor.document.lineAt(i).text.trim();
                if (!trimmed) { bodyLines.push(''); continue; }
                if (trimmed.toLowerCase().startsWith(objLower + '.')) {
                    bodyLines.push(`${innerIndent}.${trimmed.slice(obj.length + 1)}`);
                } else {
                    bodyLines.push(`${innerIndent}${trimmed}`);
                }
            }

            const newText = `${indent}With ${obj}\n${bodyLines.join('\n')}\n${indent}End With`;
            const fullRange = new vscode.Range(
                range.start.line, 0,
                range.end.line, editor.document.lineAt(range.end.line).text.length
            );
            const edit = new vscode.WorkspaceEdit();
            edit.replace(uri, fullRange, newText);
            await vscode.workspace.applyEdit(edit);
        })
    );

    // Remove Unused Variables: delete Dim declarations with no references
    context.subscriptions.push(
        vscode.commands.registerCommand('vba-runner.removeUnusedVariables', async (uri: vscode.Uri) => {
            const editor = vscode.window.activeTextEditor;
            if (!editor || editor.document.uri.toString() !== uri.toString()) return;

            const ast = lspServer.parseDocument(editor.document.getText());
            if (!ast) return;

            type DimVar = { line: number; varName: string; col: number; isSingleVar: boolean };
            const dimVars: DimVar[] = [];

            const walkBody = (body: any[]) => {
                for (const s of body) {
                    if (s.type === 'VariableDeclaration' && s.loc && !s.scope) {
                        const line0 = (s.loc.start.line as number) - 1;
                        const lineText = editor.document.lineAt(line0).text;
                        const isSingleVar = (s.declarations as any[]).length === 1;
                        for (const decl of (s.declarations as any[])) {
                            const name: string = decl.name?.name ?? '';
                            if (!name) continue;
                            const col = lineText.toLowerCase().indexOf(name.toLowerCase());
                            if (col >= 0) dimVars.push({ line: line0, varName: name, col, isSingleVar });
                        }
                    }
                    if (s.body) walkBody(s.body);
                    if (Array.isArray(s.alternate)) walkBody(s.alternate);
                    else if (s.alternate?.body) walkBody(s.alternate.body);
                    if (s.cases) for (const c of s.cases as any[]) if (c.body) walkBody(c.body);
                }
            };

            for (const stmt of (ast.body ?? [])) {
                if (stmt.type === 'ProcedureDeclaration') walkBody(stmt.body ?? []);
            }

            const toDeleteLines = new Set<number>();
            for (const { line, col, isSingleVar } of dimVars) {
                if (!isSingleVar) continue; // skip multi-var Dim to avoid removing used vars
                const refs = lspServer.getReferences(uri.toString(), line, col + 1, false);
                if (refs.length === 0) toDeleteLines.add(line);
            }

            if (toDeleteLines.size === 0) {
                vscode.window.showInformationMessage(l10n.t('No unused variables found'));
                return;
            }

            const edit = new vscode.WorkspaceEdit();
            for (const ln of [...toDeleteLines].sort((a, b) => b - a)) {
                edit.delete(uri, new vscode.Range(ln, 0, ln + 1, 0));
            }
            await vscode.workspace.applyEdit(edit);
            vscode.window.showInformationMessage(l10n.t('Deleted {0} unused variable declaration(s)', String(toDeleteLines.size)));
        })
    );

    // Organize Declarations: add Option Explicit and Dim for undeclared variables
    context.subscriptions.push(
        vscode.commands.registerCommand('vba-runner.organizeDeclarations', async (uri: vscode.Uri) => {
            const editor = vscode.window.activeTextEditor;
            if (!editor || editor.document.uri.toString() !== uri.toString()) return;

            const docText = editor.document.getText();
            const hasOptionExplicit = /^\s*Option\s+Explicit\s*$/im.test(docText);

            // Parse with Option Explicit active (prepend if missing) to detect undeclared vars
            const sourceForCheck = hasOptionExplicit ? docText : 'Option Explicit\n' + docText;
            const tokens = new Lexer(sourceForCheck).tokenize();
            const modifiedAst = new Parser(tokens, { errorRecovery: true }).parse();
            const oeResult = checkOptionExplicit(modifiedAst);

            // Also parse original AST to get correct procedure line positions
            const origAst = lspServer.parseDocument(docText);

            const edit = new vscode.WorkspaceEdit();

            if (!hasOptionExplicit) {
                edit.insert(uri, new vscode.Position(0, 0), 'Option Explicit\n');
            }

            let dimCount = 0;
            for (const stmt of (origAst?.body ?? [])) {
                if (stmt.type !== 'ProcedureDeclaration') continue;
                const procNameLower: string = (stmt as any).name?.name?.toLowerCase() ?? '';
                const undeclared = oeResult.violatedProcedures.get(procNameLower);
                if (!undeclared || undeclared.size === 0) continue;

                const procStart0 = ((stmt as any).loc.start.line as number) - 1;
                const insertLine0 = procStart0 + 1;
                const insertIndent = editor.document.lineAt(insertLine0).text.match(/^\s*/)?.[0] ?? '    ';

                const dimLines = [...undeclared.keys()]
                    .sort()
                    .map(name => `${insertIndent}Dim ${name} As Variant`)
                    .join('\n');

                edit.insert(uri, new vscode.Position(insertLine0, 0), dimLines + '\n');
                dimCount += undeclared.size;
            }

            await vscode.workspace.applyEdit(edit);

            const parts: string[] = [];
            if (!hasOptionExplicit) parts.push(l10n.t('Added Option Explicit'));
            if (dimCount > 0) parts.push(l10n.t('Added {0} Dim declaration(s)', String(dimCount)));
            if (parts.length === 0) {
                vscode.window.showInformationMessage(l10n.t('No changes'));
            } else {
                vscode.window.showInformationMessage(parts.join(', '));
            }
        })
    );

    outputChannel.appendLine('✓ Code Actions (Refactor) registered');

    outputChannel.appendLine('✓ All providers registered successfully');
    outputChannel.appendLine('📝 Open a .bas file and hover over code to test LSP features');
    console.log('🚀 VBA extension fully initialized');
}

export function deactivate() {
    console.log('VBA Runner extension deactivated');
    _hoverProviderReg?.dispose();
    _definitionProviderReg?.dispose();
    _referencesProviderReg?.dispose();
    _diagnosticCollection?.dispose();
    _testController?.dispose();
}

/**
 * Builds a TextEdit that rewrites the word just completed by the trigger
 * character `ch` to its canonical VBA keyword casing, or returns undefined when
 * there is nothing to correct.
 *
 * - For '\n', the completed word is the last identifier of the previous line.
 * - For ' '/'\t', it is the identifier ending immediately before the trigger.
 *
 * Words inside strings/comments, and member accesses (`obj.Type`) or bracketed
 * identifiers (`[Type]`), are left untouched.
 */
function keywordCasingEdit(
    document: vscode.TextDocument,
    position: vscode.Position,
    ch: string
): vscode.TextEdit | undefined {
    let lineNum: number;
    let endCol: number;
    if (ch === '\n') {
        if (position.line === 0) return undefined;
        lineNum = position.line - 1;
        endCol = document.lineAt(lineNum).text.length; // word ends at end of line
    } else {
        // The trigger char (space/tab) was inserted at position.character - 1.
        lineNum = position.line;
        endCol = position.character - 1;
    }
    if (endCol <= 0) return undefined;

    const lineText = document.lineAt(lineNum).text;
    const before = lineText.substring(0, endCol);
    const m = before.match(/([A-Za-z_][A-Za-z0-9_]*)$/);
    if (!m) return undefined;

    const word = m[1];
    const startCol = endCol - word.length;

    // Skip member access (obj.Type) and bracketed identifiers ([Type]).
    const prevCh = startCol > 0 ? lineText[startCol - 1] : '';
    if (prevCh === '.' || prevCh === '[') return undefined;

    if (isInStringOrComment(lineText, startCol)) return undefined;

    const canonical = canonicalKeyword(word);
    if (!canonical) return undefined;

    return vscode.TextEdit.replace(new vscode.Range(lineNum, startCol, lineNum, endCol), canonical);
}

