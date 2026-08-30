/**
 * Converts LSPServer diagnostics to the small subset of VS Code's diagnostic
 * API used by the extension.  Keeping this adapter free of a vscode import
 * allows the extension registration path to be tested in Node.
 */
export interface DiagnosticApi<Range, Diagnostic> {
    range(startLine: number, startCharacter: number, endLine: number, endCharacter: number): Range;
    diagnostic(range: Range, message: string, severity: number): Diagnostic;
    setSource(diagnostic: Diagnostic, source: string | undefined): void;
    setCode(diagnostic: Diagnostic, code: string | number | undefined): void;
}

export interface DiagnosticCollection<Uri, Diagnostic> {
    set(uri: Uri, diagnostics: Diagnostic[]): void;
}

export interface RawDiagnostic {
    range: { start: { line: number; character: number }; end: { line: number; character: number } };
    severity?: number;
    message: string;
    l10nKey?: string;
    l10nArgs?: string[];
    source?: string;
    code?: string | number;
}

export function publishDiagnostics<Uri, Range, Diagnostic>(
    uri: Uri,
    raw: RawDiagnostic[],
    collection: DiagnosticCollection<Uri, Diagnostic>,
    api: DiagnosticApi<Range, Diagnostic>,
    localize: (key: string, args: string[]) => string,
): void {
    const diagnostics = raw.map((item) => {
        const range = api.range(
            item.range.start.line,
            item.range.start.character,
            item.range.end.line,
            item.range.end.character,
        );
        const severity = item.severity === 1 ? 0 : item.severity === 2 ? 1 : item.severity === 3 ? 2 : 3;
        const message = item.l10nKey != null
            ? localize(item.l10nKey, item.l10nArgs ?? [])
            : item.message;
        const diagnostic = api.diagnostic(range, message, severity);
        api.setSource(diagnostic, item.source);
        api.setCode(diagnostic, item.code);
        return diagnostic;
    });
    collection.set(uri, diagnostics);
}
