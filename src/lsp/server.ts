import { Lexer } from '../engine/lexer';
import { Parser } from '../engine/parser';
import { SymbolProvider } from './symbol-provider';
import { HoverProvider } from './hover-provider';
import { DefinitionProvider } from './definition-provider';
import { CompletionProvider } from './completion-provider';
import { TestDiscovery } from './test-discovery';
import { TestRunner } from './test-runner';
import { DebugAdapter } from './debug-adapter';

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
    private testDiscovery: TestDiscovery;
    private testRunner: TestRunner;
    private debugAdapters: Map<string, DebugAdapter> = new Map();

    constructor() {
        this.symbolProvider = new SymbolProvider();
        this.hoverProvider = new HoverProvider();
        this.definitionProvider = new DefinitionProvider();
        this.completionProvider = new CompletionProvider();
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
            const tokens = new Lexer(doc.content).tokenize();
            const ast = new Parser(tokens).parse();
            return ast.diagnostics.map((d: any) => ({
                range: {
                    start: { line: d.loc.start.line - 1, character: d.loc.start.column - 1 },
                    end: { line: d.loc.end.line - 1, character: d.loc.end.column - 1 },
                },
                severity: d.severity === 'error' ? 1 : 2,
                message: d.message,
                source: 'vba-runner',
            }));
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
     * Parse document content
     */
    private parseDocument(content: string): any {
        try {
            const tokens = new Lexer(content).tokenize();
            return new Parser(tokens).parse();
        } catch (error) {
            return null;
        }
    }
}
