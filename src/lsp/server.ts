import { Lexer } from '../engine/lexer';
import { Parser } from '../engine/parser';
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
            const tokens = new Lexer(this.stripClsHeader(doc.content)).tokenize();
            const ast = new Parser(tokens).parse();
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
            return [...parseDiags, ...deadCodeWarnings];
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
     * Parse document content
     */
    parseDocument(content: string): any {
        try {
            const tokens = new Lexer(this.stripClsHeader(content)).tokenize();
            return new Parser(tokens).parse();
        } catch (error) {
            return null;
        }
    }
}
