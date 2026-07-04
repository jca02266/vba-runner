import fs from 'node:fs';
import path from 'node:path';
import { Lexer } from '../engine/lexer';
import { Parser } from '../engine/parser';
import { detectRangeAccess } from '../engine/range-access-detector';
import { lintProgram, findLoopContinueLabels } from '../engine/vba-lint';
import { checkOptionExplicit } from '../engine/option-explicit-checker';
import { stripVBAFileHeader } from '../engine/preprocessor';
import { Statement, GoToStatement } from '../engine/parser';
import { findLabelDefinition, findGoToReferences, isOnLabel } from './label-navigator';
import { inferProcedureHints, buildProcMap } from './variant-type-inferencer';
import { analyzeDefUse } from '../engine/def-use-analyzer';
import { ProcedureDeclaration } from '../engine/parser';
import { SymbolProvider } from './symbol-provider';
import { HoverProvider } from './hover-provider';
import { DefinitionProvider } from './definition-provider';
import { CompletionProvider, generateDefaultTypeStubsJson, parseTypeStubsJson } from './completion-provider';
import { checkUnknownTypes, collectUserDefinedTypeNames } from './unknown-type-checker';
import { findAllReferences, LocationInfo } from './references-provider';
import { buildScopedSymbolTable, getWordAtPosition } from './symbol-table';
import { RenameProvider } from './rename-provider';
import { CodeLensProvider, TestRunResult } from './code-lens-provider';
import { CallGraphProvider, CallGraph } from './call-graph-provider';
import { TestDiscovery } from './test-discovery';
import { TestRunner } from './test-runner';
import { DebugAdapter } from './debug-adapter';
import { FoldingRangeProvider, FoldingRange } from './folding-range-provider';
import { SignatureHelpProvider, SignatureHelpResult } from './signature-help-provider';

export interface TextDocument {
    uri: string;
    content: string;
    version: number;
}

export interface LSPServerConfig {
    capabilities: {
        textDocumentSync: number;
        completionProvider: any;
        hoverProvider: boolean;
        definitionProvider: boolean;
        documentSymbolProvider: boolean;
        signatureHelpProvider?: { triggerCharacters: string[] };
    };
}

export class LSPServer {
    private documents: Map<string, TextDocument> = new Map();
    private documentVersions: Map<string, number> = new Map();
    /** ディスク上の VBA ファイルのキャッシュ（遅延スキャンで追加） */
    private workspaceDocuments: Map<string, TextDocument> = new Map();
    /** スキャン済みディレクトリ（同じディレクトリを2回読まないため） */
    private scannedDirectories: Set<string> = new Set();
    private symbolProvider: SymbolProvider;
    private hoverProvider: HoverProvider;
    private definitionProvider: DefinitionProvider;
    private completionProvider: CompletionProvider;
    private renameProvider: RenameProvider;
    private codeLensProvider: CodeLensProvider;
    private callGraphProvider: CallGraphProvider;
    private foldingRangeProvider: FoldingRangeProvider;
    private testDiscovery: TestDiscovery;
    private testRunner: TestRunner;
    private signatureHelpProvider: SignatureHelpProvider;
    private debugAdapters: Map<string, DebugAdapter> = new Map();
    private testResultCache = new Map<string, Map<string, TestRunResult>>();

    constructor() {
        this.symbolProvider = new SymbolProvider();
        this.hoverProvider = new HoverProvider();
        this.definitionProvider = new DefinitionProvider();
        this.completionProvider = new CompletionProvider();
        this.renameProvider = new RenameProvider();
        this.codeLensProvider = new CodeLensProvider();
        this.callGraphProvider = new CallGraphProvider();
        this.foldingRangeProvider = new FoldingRangeProvider();
        this.testDiscovery = new TestDiscovery();
        this.testRunner = new TestRunner();
        this.signatureHelpProvider = new SignatureHelpProvider();
    }

    setDebugPrintHandler(onPrint: (s: string) => void): void {
        this.testRunner = new TestRunner(onPrint);
    }

    setTestResults(uri: string, results: Array<{name: string; state: string; duration: number; message?: string}>): void {
        const map = new Map<string, TestRunResult>();
        for (const r of results) {
            map.set(r.name.toLowerCase(), { state: r.state as TestRunResult['state'], duration: r.duration, message: r.message });
        }
        this.testResultCache.set(uri, map);
    }

    /**
     * Initialize the server with client capabilities
     */
    initialize(): LSPServerConfig {
        return {
            capabilities: {
                textDocumentSync: 1, // FULL
                completionProvider: {
                    resolveProvider: false,
                    triggerCharacters: ['.'],
                },
                hoverProvider: true,
                definitionProvider: true,
                documentSymbolProvider: true,
                signatureHelpProvider: {
                    triggerCharacters: ['(', ','],
                },
            },
        };
    }

    /**
     * Check whether a document is currently open
     */
    hasDocument(uri: string): boolean {
        return this.documents.has(uri);
    }

    /**
     * ワークスペース初期スキャンでディスク上のファイルを登録する。
     * エディターで開いているファイルより優先度は低い。
     */
    loadWorkspaceFile(uri: string, content: string): void {
        this.workspaceDocuments.set(uri, { uri, content, version: 0 });
    }

    /**
     * ワークスペースからファイルが削除されたときに呼ぶ。
     * ディレクトリのスキャン済みフラグも無効化し、次回 F12 時に再スキャンされるようにする。
     */
    unloadWorkspaceFile(uri: string): void {
        this.workspaceDocuments.delete(uri);
        try {
            const dir = path.dirname(uriToPath(uri));
            this.scannedDirectories.delete(dir);
        } catch { /* file URI 以外は無視 */ }
    }

    /**
     * F12/Shift+F12 の先頭で呼ぶ。uri と同じディレクトリを初回のみスキャンして
     * workspaceDocuments に追加する。同じディレクトリは2回スキャンしない。
     */
    private ensureDirectoryScanned(uri: string): void {
        let dir: string;
        try {
            dir = path.dirname(uriToPath(uri));
        } catch {
            return; // file:// 以外の URI（テスト環境等）は無視
        }

        if (this.scannedDirectories.has(dir)) return;
        this.scannedDirectories.add(dir);

        let entries: string[];
        try {
            entries = fs.readdirSync(dir);
        } catch {
            return; // ディレクトリが存在しない場合は無視
        }

        for (const entry of entries) {
            const ext = path.extname(entry).toLowerCase();
            if (ext !== '.bas' && ext !== '.cls' && ext !== '.frm') continue;

            const fullPath = path.join(dir, entry);
            const entryUri = pathToUri(fullPath);

            if (this.workspaceDocuments.has(entryUri) || this.documents.has(entryUri)) continue;

            try {
                const content = fs.readFileSync(fullPath, 'utf-8');
                this.workspaceDocuments.set(entryUri, { uri: entryUri, content, version: 0 });
            } catch { /* 読み取り不可は無視 */ }
        }
    }

    /**
     * クロスファイル検索に使う全ドキュメントのビュー。
     * エディターで開いているファイル（documents）がディスクキャッシュ（workspaceDocuments）を上書きする。
     */
    private allDocuments(): Map<string, TextDocument> {
        const merged = new Map(this.workspaceDocuments);
        for (const [uri, doc] of this.documents) {
            merged.set(uri, doc);
        }
        return merged;
    }

    /**
     * Open a document
     */
    didOpen(uri: string, content: string): void {
        const version = (this.documentVersions.get(uri) ?? 0) + 1;
        this.documents.set(uri, { uri, content, version });
        this.documentVersions.set(uri, version);
        // workspaceDocuments にない新規ファイル（ワークスペース外や新規作成）も登録しておく
        if (!this.workspaceDocuments.has(uri)) {
            this.workspaceDocuments.set(uri, { uri, content, version });
        }
    }

    /**
     * Change document content
     */
    didChange(uri: string, content: string): void {
        const version = (this.documentVersions.get(uri) ?? 0) + 1;
        this.documents.set(uri, { uri, content, version });
        this.documentVersions.set(uri, version);
        // ワークスペースキャッシュも最新の編集内容に追随させる
        this.workspaceDocuments.set(uri, { uri, content, version });
    }

    /**
     * Close a document
     */
    didClose(uri: string): void {
        this.documents.delete(uri);
        this.debugAdapters.delete(uri);
        // workspaceDocuments は残す（ディスク上のファイルは存在し続ける）
    }

    /**
     * Get document symbols (outline)
     */
    getDocumentSymbols(uri: string): any[] {
        const doc = this.documents.get(uri);
        if (!doc) return [];

        const ast = this.parseDocument(doc.content);
        if (!ast) return [];

        this.symbolProvider.setDocumentUri(uri);
        return this.symbolProvider.extractSymbols(ast.body, doc.content);
    }

    getFoldingRanges(uri: string): FoldingRange[] {
        const doc = this.documents.get(uri);
        if (!doc) return [];

        const ast = this.parseDocument(doc.content);
        if (!ast) return [];

        return this.foldingRangeProvider.getFoldingRanges(ast.body);
    }

    /**
     * Get hover information
     */
    getHover(uri: string, line: number, character: number): any {
        const doc = this.documents.get(uri) ?? this.workspaceDocuments.get(uri);
        if (!doc) return null;

        const ast = this.parseDocument(doc.content);
        if (!ast) return null;

        // メンバーホバー: obj.Member にカーソルがある場合はシグネチャを表示する
        const memberInfo = this.completionProvider.getMemberHoverInfo(doc.content, line, character, ast.body);
        const symbolHover = this.hoverProvider.getHoverInfo(ast.body, doc.content, line, character);

        if (memberInfo) {
            const memberContents = `\`\`\`vb\n${memberInfo.detail}\n\`\`\``;
            if (symbolHover) {
                return { ...symbolHover, contents: `${memberContents}\n\n${symbolHover.contents}` };
            }
            const lines = doc.content.split('\n');
            const currentLine = lines[line] ?? '';
            const col = currentLine.substring(0, character).match(/([a-zA-Z_][a-zA-Z0-9_]*)$/)?.[1].length ?? 0;
            return {
                contents: memberContents,
                range: {
                    start: { line, character: character - col },
                    end:   { line, character: character - col + memberInfo.label.length },
                },
            };
        }

        return symbolHover;
    }

    /**
     * Get definition location
     */
    getDefinition(uri: string, line: number, character: number): any {
        const doc = this.documents.get(uri);
        if (!doc) return null;

        const ast = this.parseDocument(doc.content);
        if (!ast) return null;

        // ラベル定義ジャンプ（GoTo のラベル名 → LabelStatement）
        const labelDef = findLabelDefinition(ast.body, line, character, uri);
        if (labelDef) return labelDef;

        // クロスファイル検索の前に同ディレクトリを遅延スキャン
        this.ensureDirectoryScanned(uri);

        this.definitionProvider.setDocumentUri(uri);
        try {
            const result = this.definitionProvider.getDefinition(ast.body, doc.content, line, character);
            if (result) return result;
        } catch {
            return null;
        }

        // 現在ファイルで見つからなければ他のドキュメントのモジュールレベルシンボルを検索
        const word = getWordAtPosition(doc.content, line, character);
        if (!word) return null;
        const wordLower = word.toLowerCase();

        for (const [otherUri, otherDoc] of this.allDocuments()) {
            if (otherUri === uri) continue;
            const otherAst = this.parseDocument(otherDoc.content);
            if (!otherAst) continue;
            const otherTable = buildScopedSymbolTable(otherAst.body);
            const entry = otherTable.moduleSymbols.get(wordLower);
            if (entry) {
                return { uri: otherUri, range: entry.range };
            }
        }

        return null;
    }

    /**
     * Get all references to the symbol at the given position
     */
    getReferences(uri: string, line: number, character: number, includeDeclaration: boolean): any[] {
        const doc = this.documents.get(uri);
        if (!doc) return [];

        const ast = this.parseDocument(doc.content);
        if (!ast) return [];

        // ラベル参照ジャンプ（LabelStatement → GoTo 一覧）
        if (isOnLabel(ast.body, line, character)) {
            return findGoToReferences(ast.body, line, character, uri, includeDeclaration);
        }

        // クロスファイル検索の前に同ディレクトリを遅延スキャン
        this.ensureDirectoryScanned(uri);

        const word = getWordAtPosition(doc.content, line, character);
        if (!word) return [];

        // ローカルシンボル（プロシージャ内ローカル変数）は現在ファイルのみ検索
        const table = buildScopedSymbolTable(ast.body);
        const wordLower = word.toLowerCase();
        const isOnlyLocal = !table.moduleSymbols.has(wordLower)
            && table.procedures.some(p => p.localSymbols.has(wordLower));

        if (isOnlyLocal) {
            return findAllReferences(doc.content, word, uri, ast.body, includeDeclaration, line);
        }

        // モジュールレベルシンボルはワークスペース全体を横断検索
        const allRefs: LocationInfo[] = [];
        for (const [docUri, docDoc] of this.allDocuments()) {
            const docAst = this.parseDocument(docDoc.content);
            if (!docAst) continue;
            const refs = findAllReferences(docDoc.content, word, docUri, docAst.body, includeDeclaration);
            allRefs.push(...refs);
        }
        return allRefs;
    }

    /**
     * Get workspace edits for renaming the symbol at the given position
     */
    getRename(uri: string, line: number, character: number, newName: string): any[] | null {
        const doc = this.documents.get(uri);
        if (!doc) return null;

        const ast = this.parseDocument(doc.content);
        if (!ast) return null;

        this.renameProvider.setDocumentUri(uri);
        return this.renameProvider.getRename(ast.body, doc.content, line, character, newName);
    }

    /**
     * Get code lens items for document
     */
    getCodeLens(uri: string): any[] {
        const doc = this.documents.get(uri);
        if (!doc) return [];

        const ast = this.parseDocument(doc.content);
        if (!ast) return [];

        return this.codeLensProvider.getCodeLens(ast.body, doc.content, uri, this.testResultCache.get(uri));
    }

    /**
     * Get completion suggestions
     */
    getCompletions(uri: string, line: number, character: number): any[] {
        const doc = this.documents.get(uri);
        if (!doc) return [];

        const ast = this.parseDocument(doc.content);
        if (!ast) return [];

        // 複数モジュールの AST を統合してクロスモジュール補完を実現する
        const allStatements: Statement[] = [...ast.body];
        for (const [otherUri, otherDoc] of this.allDocuments()) {
            if (otherUri === uri) continue;
            try {
                const otherAst = this.parseDocument(otherDoc.content);
                if (otherAst?.body) allStatements.push(...otherAst.body);
            } catch { /* 壊れたドキュメントはスキップ */ }
        }

        return this.completionProvider.getCompletions(allStatements, doc.content, line, character);
    }

    /** 外部型定義ファイル (vba-types.json) を読み込んで補完プロバイダーに設定する。 */
    loadTypeStubs(workspaceRoot: string): void {
        const stubPath = path.join(workspaceRoot, 'vba-types.json');
        if (!fs.existsSync(stubPath)) return;
        try {
            const raw = fs.readFileSync(stubPath, 'utf-8');
            this.completionProvider.setTypeStubs(parseTypeStubsJson(raw));
        } catch (e) {
            console.error('[VBA] Failed to load vba-types.json:', e);
        }
    }

    /** 外部型定義の更新（ファイルウォッチャー用）。 */
    reloadTypeStubs(workspaceRoot: string): void {
        this.loadTypeStubs(workspaceRoot);
    }

    /** BUILTIN_MEMBERS の内容を vba-types.json 形式の JSON 文字列として返す。 */
    generateDefaultTypeStubsJson(): string {
        return generateDefaultTypeStubsJson();
    }

    /** 全ワークスペースドキュメントから Class / Type / Enum の型名を収集する。 */
    private collectAllUserDefinedTypeNames(): Set<string> {
        const names = new Set<string>();
        for (const [, doc] of this.allDocuments()) {
            try {
                const ast = this.parseDocument(doc.content);
                if (ast?.body) {
                    for (const name of collectUserDefinedTypeNames(ast.body)) names.add(name);
                }
            } catch { /* ignore */ }
        }
        return names;
    }

    /**
     * Get signature help for a function call at the given position
     */
    getSignatureHelp(uri: string, line: number, character: number): SignatureHelpResult | null {
        const doc = this.documents.get(uri) ?? this.workspaceDocuments.get(uri);
        if (!doc) return null;

        const ast = this.parseDocument(doc.content);
        if (!ast) return null;

        return this.signatureHelpProvider.getSignatureHelp(ast.body, doc.content, line, character);
    }

    /**
     * Discover tests in document
     */
    discoverTests(uri: string): any[] {
        const doc = this.documents.get(uri);
        if (!doc) return [];

        const ast = this.parseDocument(doc.content);
        if (!ast) return [];

        return this.testDiscovery.discoverTests(ast.body);
    }

    /**
     * Run tests in document
     */
    runTests(uri: string): any[] {
        const doc = this.documents.get(uri);
        if (!doc) return [];

        const ast = this.parseDocument(doc.content);
        if (!ast) return [];

        return this.testRunner.runTests(ast.body);
    }

    /**
     * Get diagnostics (parse errors/warnings) for document
     */
    getDiagnostics(uri: string): any[] {
        const doc = this.documents.get(uri);
        if (!doc) return [];

        try {
            const lexer = new Lexer(stripVBAFileHeader(doc.content));
            const tokens = lexer.tokenize();
            const lexerDiags = lexer.diagnostics.map(d => ({
                range: {
                    start: { line: d.line - 1, character: d.column - 1 },
                    end:   { line: d.line - 1, character: d.column - 1 + 1 },
                },
                severity: 1, // Error
                message: d.message,
                source: 'vba-runner',
            }));
            const ast = new Parser(tokens, { errorRecovery: true }).parse();
            checkOptionExplicit(ast); // populate ast.diagnostics with undeclared variable errors
            const parseDiags = ast.diagnostics.map((d: any) => ({
                range: {
                    start: { line: d.loc.start.line - 1, character: d.loc.start.column - 1 },
                    end: { line: d.loc.end.line - 1, character: d.loc.end.column - 1 },
                },
                severity: d.severity === 'error' ? 1 : 2,
                message: d.message,
                source: 'vba-runner',
            }));
            const deadCodeWarnings = this.codeLensProvider.getDeadCodeWarnings(ast.body, doc.content, uri);
            const rangeAccessHints = detectRangeAccess(ast).map(hit => ({
                range: {
                    start: { line: hit.line, character: hit.column },
                    end:   { line: hit.line, character: hit.column + hit.varName.length },
                },
                severity: 4, // Hint
                code: 'VBA011',
                message: hit.message,
                l10nKey: hit.l10nKey,
                l10nArgs: hit.l10nArgs,
                source: 'vba-dataflow',
            }));
            const vbaLintDiags = lintProgram(ast).map(d => ({
                range: {
                    start: { line: d.line, character: d.column },
                    end:   { line: d.endLine, character: d.endColumn },
                },
                severity: d.severity,
                code: d.code,
                message: d.message,
                l10nKey: d.l10nKey,
                l10nArgs: d.l10nArgs,
                source: `vba-lint(${d.code})`,
            }));
            // VBA016: 未知の型名チェック
            const knownTypes = this.completionProvider.getKnownTypeNamesForDiagnostics();
            for (const name of this.collectAllUserDefinedTypeNames()) knownTypes.add(name);
            const unknownTypeDiags = checkUnknownTypes(ast.body, knownTypes).map(d => ({
                range: {
                    start: { line: d.line, character: d.column },
                    end:   { line: d.endLine, character: d.endColumn },
                },
                severity: d.severity,
                code: d.code,
                message: d.message,
                l10nKey: d.l10nKey,
                l10nArgs: d.l10nArgs,
                source: `vba-lint(${d.code})`,
            }));

            return [...lexerDiags, ...parseDiags, ...deadCodeWarnings, ...rangeAccessHints, ...vbaLintDiags, ...unknownTypeDiags];
        } catch {
            return [];
        }
    }

    /**
     * Create debug adapter for document
     */
    createDebugAdapter(uri: string): DebugAdapter | null {
        const doc = this.documents.get(uri);
        if (!doc) return null;

        // Derive module name from URI (filename without extension)
        const uriPath = uri.startsWith('file://') ? decodeURIComponent(uri.slice(7)) : uri;
        const moduleName = uriPath.split('/').pop()?.replace(/\.[^.]+$/, '') ?? 'Module1';

        const adapter = new DebugAdapter(doc.content, moduleName, uriPath);
        this.debugAdapters.set(uri, adapter);
        return adapter;
    }

    /**
     * Get debug adapter for document
     */
    getDebugAdapter(uri: string): DebugAdapter | null {
        return this.debugAdapters.get(uri) ?? null;
    }

    /**
     * Strip .cls file header (VERSION 1.0 CLASS / BEGIN...END block).
     * Replaces header lines with empty lines to preserve line numbers.
     */
    getCallGraph(_currentUri: string): CallGraph {
        const fileMap = new Map<string, { statements: any[], uri: string }>();
        for (const [uri, doc] of this.documents) {
            const ast = this.parseDocument(doc.content);
            if (ast?.body) {
                fileMap.set(uri, { statements: ast.body, uri });
            }
        }
        return this.callGraphProvider.buildCallGraph(fileMap);
    }

    buildCallGraphFromFiles(fileContents: Map<string, string>): CallGraph {
        const fileMap = new Map<string, { statements: any[], uri: string }>();
        for (const [uri, content] of fileContents) {
            const ast = this.parseDocument(content);
            if (ast?.body) {
                fileMap.set(uri, { statements: ast.body, uri });
            }
        }
        return this.callGraphProvider.buildCallGraph(fileMap);
    }

    /**
     * Get code actions for a selection range.
     * Offers "Extract Function" when the selection is within a procedure.
     */
    getCodeActions(
        uri: string,
        rawRange: { start: { line: number; character: number }; end: { line: number; character: number } },
    ): any[] {
        const doc = this.documents.get(uri);
        if (!doc) return [];

        // VS Code の行ドラッグ選択では range.end が「選択した最後の行の次の行・
        // character 0」になることが多い。この場合、実質的な選択末尾は1行前。
        // 正規化した range を以降の解析・呼び出し元への返却双方で一貫して使う。
        const range = (rawRange.end.character === 0 && rawRange.end.line > rawRange.start.line)
            ? { start: rawRange.start, end: { line: rawRange.end.line - 1, character: 0 } }
            : rawRange;

        try {
            const tokens    = new Lexer(stripVBAFileHeader(doc.content)).tokenize();
            const ast       = new Parser(tokens, { errorRecovery: true }).parse();
            const startLine = range.start.line + 1;  // 0-based → 1-based
            const endLine   = range.end.line   + 1;

            // 選択範囲を含むプロシージャを探す
            const proc = ast.body.find((s: any) =>
                s.type === 'ProcedureDeclaration' &&
                s.loc  != null                    &&
                s.loc.start.line <= startLine     &&
                s.loc.end.line   >= endLine,
            ) as ProcedureDeclaration | undefined;
            if (!proc) return [];

            const result = analyzeDefUse(proc, startLine, endLine);
            if (
                result.inputs.length  === 0 &&
                result.outputs.length === 0 &&
                result.locals.length  === 0
            ) return [];

            const inputParams  = result.inputs.map(v  => `ByVal ${v} As Variant`);
            const outputParams = result.outputs.map(v => `ByRef ${v} As Variant`);
            const allParams    = [...inputParams, ...outputParams].join(', ');
            const localDecls   = result.locals.map(v => `    Dim ${v} As Variant`).join('\n');
            const callArgs     = [...result.inputs, ...result.outputs].join(', ');
            const newProcName  = 'ExtractedSub';
            const todoComment  = `    ' TODO: Move extracted code here`;
            const procBody     = localDecls ? `${localDecls}\n${todoComment}` : todoComment;
            const procSignature = `Private Sub ${newProcName}(${allParams})\n${procBody}\nEnd Sub`;
            const callStatement = `${newProcName}(${callArgs})`;

            return [{
                title: `⚡ Extract Function: ${callStatement}`,
                kind: 'refactor.extract.function',
                command: {
                    title: 'Extract Function',
                    command: 'vba-runner.doExtractFunction',
                    arguments: [uri, range, result, procSignature, callStatement],
                },
            }];
        } catch {
            return [];
        }
    }

    /**
     * Build the text edits for extracting a selected range (already normalized by
     * getCodeActions) into a new procedure. Pure text computation — no vscode
     * dependency — so extension.ts only needs to apply the returned edits.
     */
    buildExtractFunctionEdit(
        uri: string,
        range: { start: { line: number; character: number }; end: { line: number; character: number } },
        procName: string,
        result: { inputs: string[]; outputs: string[]; locals: string[] },
        procSignature: string,
        callStatement: string,
    ): {
        replaceRange: { startLine: number; endLine: number; endCharacter: number };
        replaceText: string;
        insertLine: number;
        insertText: string;
    } | null {
        const doc = this.documents.get(uri);
        if (!doc) return null;

        const lines     = doc.content.split('\n');
        const startLine = range.start.line;
        const endLine   = range.end.line;

        const selectedLines = lines.slice(startLine, endLine + 1);

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

        const newProcText = [finalSigLine, ...extraDims, ...reindented, 'End Sub'].join('\n');

        // Find containing procedure's last line (0-based) for insertion point
        const ast = this.parseDocument(doc.content);
        let procEndLine = endLine;
        if (ast?.body) {
            for (const stmt of ast.body as any[]) {
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

        const callIndent  = selectedLines[0]?.match(/^\s*/)?.[0] ?? '';
        const endLineText = lines[endLine] ?? '';

        return {
            replaceRange: { startLine, endLine, endCharacter: endLineText.length },
            replaceText: callIndent + finalCall,
            insertLine: procEndLine + 1,
            insertText: '\n' + newProcText + '\n',
        };
    }

    /**
     * Get inlay hints for Variant-typed variables in all procedures of the document.
     * A single shared memo prevents redundant function-return-type resolution.
     */
    getVariantTypeHints(uri: string): Array<{ line: number; character: number; label: string }> {
        const doc = this.documents.get(uri);
        if (!doc) return [];
        const ast = this.parseDocument(doc.content);
        if (!ast) return [];

        const allProcs = buildProcMap(ast.body);
        const memo = new Map();
        return [...allProcs.values()]
            .flatMap(proc => inferProcedureHints(proc, allProcs, memo))
            .filter(h => h.inferredType !== null)
            .map(h => ({
                line:      h.line - 1,
                character: h.endColumn - 1,
                label:     ` As ${h.inferredType}`,
            }));
    }

    /**
     * Get inlay hints for loop-continue GoTo statements
     */
    getInlayHints(uri: string): Array<{ line: number; character: number; label: string }> {
        const doc = this.documents.get(uri);
        if (!doc) return [];

        const ast = this.parseDocument(doc.content);
        if (!ast) return [];

        const hints: Array<{ line: number; character: number; label: string }> = [];
        collectGoToContinueHints(ast.body, new Set(), hints);
        return hints;
    }

    /**
     * Parse document content
     */
    parseDocument(content: string): any {
        try {
            const tokens = new Lexer(stripVBAFileHeader(content)).tokenize();
            return new Parser(tokens, { errorRecovery: true }).parse();
        } catch (error) {
            return null;
        }
    }
}

function collectGoToContinueHints(
    stmts: Statement[],
    continueLabels: Set<string>,
    out: Array<{ line: number; character: number; label: string }>,
): void {
    for (const stmt of stmts) {
        if (stmt.type === 'GoToStatement') {
            const gs = stmt as GoToStatement;
            if (continueLabels.has(gs.label.toLowerCase())) {
                const loc = (gs as any).loc;
                if (loc) {
                    out.push({
                        line: loc.end.line - 1,
                        character: loc.end.column - 1,
                        label: ' ⟨loop continue⟩',
                    });
                }
            }
        } else if (stmt.type === 'ProcedureDeclaration') {
            // プロシージャに入るたびにそのbodyからcontinueLabelsを再計算する
            const procBody: Statement[] = (stmt as any).body ?? [];
            const procLabels = findLoopContinueLabels(procBody);
            collectGoToContinueHints(procBody, procLabels, out);
        } else if (stmt.type === 'ClassDeclaration') {
            collectGoToContinueHints((stmt as any).procedures ?? [], continueLabels, out);
        } else if (stmt.type === 'IfStatement') {
            const is = stmt as any;
            const cons = Array.isArray(is.consequent) ? is.consequent : (is.consequent ? [is.consequent] : []);
            const alt  = Array.isArray(is.alternate)  ? is.alternate  : (is.alternate  ? [is.alternate]  : []);
            collectGoToContinueHints(cons, continueLabels, out);
            collectGoToContinueHints(alt,  continueLabels, out);
        } else if (stmt.type === 'SelectCaseStatement') {
            for (const clause of (stmt as any).cases ?? []) {
                collectGoToContinueHints(clause.body ?? [], continueLabels, out);
            }
            collectGoToContinueHints((stmt as any).elseBody ?? [], continueLabels, out);
        } else {
            // For/ForEach/DoWhile/While/With など body を持つ汎用ケース
            const body: Statement[] = (stmt as any).body ?? [];
            if (body.length) collectGoToContinueHints(body, continueLabels, out);
        }
    }
}

// ─── URI / パス変換ユーティリティ ───────────────────────────────────────────

function uriToPath(uri: string): string {
    const url = new URL(uri);
    let p = decodeURIComponent(url.pathname);
    // Windows: /C:/path → C:/path
    if (/^\/[A-Za-z]:/.test(p)) p = p.slice(1);
    return p;
}

function pathToUri(filePath: string): string {
    const normalized = filePath.replace(/\\/g, '/');
    return normalized.startsWith('/') ? `file://${normalized}` : `file:///${normalized}`;
}
