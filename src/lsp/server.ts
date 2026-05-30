import { Lexer } from '../engine/lexer';
import { Parser } from '../engine/parser';
import { detectRangeAccess } from '../engine/range-access-detector';
import { lintProgram, findLoopContinueLabels } from '../engine/vba-lint';
import { Statement, GoToStatement } from '../engine/parser';
import { findLabelDefinition, findGoToReferences, isOnLabel } from './label-navigator';
import { inferProcedureHints, buildProcMap } from './variant-type-inferencer';
import { analyzeDefUse } from '../engine/def-use-analyzer';
import { ProcedureDeclaration } from '../engine/parser';
import { SymbolProvider } from './symbol-provider';
import { HoverProvider } from './hover-provider';
import { DefinitionProvider } from './definition-provider';
import { CompletionProvider } from './completion-provider';
import { ReferencesProvider } from './references-provider';
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
    private symbolProvider: SymbolProvider;
    private hoverProvider: HoverProvider;
    private definitionProvider: DefinitionProvider;
    private completionProvider: CompletionProvider;
    private referencesProvider: ReferencesProvider;
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
        this.referencesProvider = new ReferencesProvider();
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
     * Open a document
     */
    didOpen(uri: string, content: string): void {
        const version = (this.documentVersions.get(uri) ?? 0) + 1;
        this.documents.set(uri, { uri, content, version });
        this.documentVersions.set(uri, version);
    }

    /**
     * Change document content
     */
    didChange(uri: string, content: string): void {
        const version = (this.documentVersions.get(uri) ?? 0) + 1;
        this.documents.set(uri, { uri, content, version });
        this.documentVersions.set(uri, version);
    }

    /**
     * Close a document
     */
    didClose(uri: string): void {
        this.documents.delete(uri);
        this.debugAdapters.delete(uri);
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

        this.definitionProvider.setDocumentUri(uri);
        try {
            return this.definitionProvider.getDefinition(ast.body, doc.content, line, character);
        } catch {
            return null;
        }
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

        this.referencesProvider.setDocumentUri(uri);
        return this.referencesProvider.getReferences(ast.body, doc.content, line, character, includeDeclaration);
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
            const lexer = new Lexer(this.stripClsHeader(doc.content));
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
                message: hit.kind === 'index-call'
                    ? `Excel依存：Range 変数 '${hit.varName}' への添字アクセス（${hit.varName}(...) は ${hit.varName}.Item(...) と等価）`
                    : hit.kind === 'member-call'
                        ? `Excel依存：Range 変数 '${hit.varName}' へのメソッド呼び出し: .${hit.property}()`
                        : `Excel依存：Range 変数 '${hit.varName}' へのプロパティアクセス: .${hit.property}`,
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

        const ast = this.parseDocument(doc.content);
        if (!ast) return null;

        const adapter = new DebugAdapter(ast);
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
            const tokens    = new Lexer(this.stripClsHeader(doc.content)).tokenize();
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
            const procBody     = localDecls ? `${localDecls}\n    ' TODO: 抽出したコードをここに移動` : `    ' TODO: 抽出したコードをここに移動`;
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

    private stripClsHeader(content: string): string {
        const lines = content.split('\n');
        if (!lines[0]?.trimEnd().toUpperCase().startsWith('VERSION')) return content;
        const result = [...lines];
        let i = 0;
        result[i] = '';
        i++;
        while (i < result.length) {
            const trimmed = result[i].trimEnd().toUpperCase();
            result[i] = '';
            i++;
            if (trimmed === 'END') break;
        }
        return result.join('\n');
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
            const tokens = new Lexer(this.stripClsHeader(content)).tokenize();
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
