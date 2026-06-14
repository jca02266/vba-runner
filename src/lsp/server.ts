import fs from 'node:fs';
import path from 'node:path';
import { Lexer } from '../engine/lexer';
import { Parser } from '../engine/parser';
import { detectRangeAccess } from '../engine/range-access-detector';
import { lintProgram, findLoopContinueLabels } from '../engine/vba-lint';
import { stripVBAFileHeader } from '../engine/preprocessor';
import { Statement, GoToStatement } from '../engine/parser';
import { findLabelDefinition, findGoToReferences, isOnLabel } from './label-navigator';
import { inferProcedureHints, buildProcMap } from './variant-type-inferencer';
import { analyzeDefUse } from '../engine/def-use-analyzer';
import { ProcedureDeclaration } from '../engine/parser';
import { SymbolProvider } from './symbol-provider';
import { HoverProvider } from './hover-provider';
import { DefinitionProvider } from './definition-provider';
import { CompletionProvider } from './completion-provider';
import { findAllReferences, LocationInfo } from './references-provider';
import { buildScopedSymbolTable, getWordAtPosition } from './symbol-table';
import { RenameProvider } from './rename-provider';
import { CodeLensProvider } from './code-lens-provider';
import { CallGraphProvider, CallGraph } from './call-graph-provider';
import { TestDiscovery } from './test-discovery';
import { TestRunner } from './test-runner';
import { DebugAdapter } from './debug-adapter';
import { FoldingRangeProvider, FoldingRange } from './folding-range-provider';

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
    private debugAdapters: Map<string, DebugAdapter> = new Map();

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
        return this.symbolProvider.extractSymbols(ast.body);
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
        const doc = this.documents.get(uri);
        if (!doc) return null;

        const ast = this.parseDocument(doc.content);
        if (!ast) return null;

        return this.hoverProvider.getHoverInfo(ast.body, doc.content, line, character);
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

        return this.codeLensProvider.getCodeLens(ast.body, doc.content, uri);
    }

    /**
     * Get completion suggestions
     */
    getCompletions(uri: string, line: number, character: number): any[] {
        const doc = this.documents.get(uri);
        if (!doc) return [];

        const ast = this.parseDocument(doc.content);
        if (!ast) return [];

        return this.completionProvider.getCompletions(ast.body, doc.content, line, character);
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
                ...(hit.kind === 'index-call' ? {
                    message: `Excel dependency: subscript access on Range variable '${hit.varName}' (${hit.varName}(...) is equivalent to ${hit.varName}.Item(...))`,
                    l10nKey: "Excel dependency: subscript access on Range variable '{0}' ({0}(...) is equivalent to {0}.Item(...))",
                    l10nArgs: [hit.varName, hit.varName, hit.varName],
                } : hit.kind === 'member-call' ? {
                    message: `Excel dependency: method call on Range variable '${hit.varName}': .${hit.property ?? ''}()`,
                    l10nKey: "Excel dependency: method call on Range variable '{0}': .{1}()",
                    l10nArgs: [hit.varName, hit.property ?? ''],
                } : {
                    message: `Excel dependency: property access on Range variable '${hit.varName}': .${hit.property ?? ''}`,
                    l10nKey: "Excel dependency: property access on Range variable '{0}': .{1}",
                    l10nArgs: [hit.varName, hit.property ?? ''],
                }),
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
            return [...lexerDiags, ...parseDiags, ...deadCodeWarnings, ...rangeAccessHints, ...vbaLintDiags];
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
        range: { start: { line: number; character: number }; end: { line: number; character: number } },
    ): any[] {
        const doc = this.documents.get(uri);
        if (!doc) return [];

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
