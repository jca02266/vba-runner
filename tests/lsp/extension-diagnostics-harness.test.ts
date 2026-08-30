import { evalVBASingle, assert } from '../../test-libs/test-runner';
import { LSPServer } from '../../src/lsp/server';
import { publishDiagnostics, DiagnosticApi } from '../../src/lsp/diagnostic-publisher';

type FakeRange = { start: { line: number; character: number }; end: { line: number; character: number } };
type FakeDiagnostic = FakeRange & { message: string; severity: number; source?: string; code?: string | number };

class FakeDiagnosticCollection {
    private readonly entries = new Map<string, FakeDiagnostic[]>();

    set(uri: string, diagnostics: FakeDiagnostic[]): void {
        this.entries.set(uri, diagnostics);
    }

    get(uri: string): FakeDiagnostic[] {
        return this.entries.get(uri) ?? [];
    }
}

const fakeApi: DiagnosticApi<FakeRange, FakeDiagnostic> = {
    range: (startLine, startCharacter, endLine, endCharacter) => ({
        start: { line: startLine, character: startCharacter },
        end: { line: endLine, character: endCharacter },
    }),
    diagnostic: (range, message, severity) => ({ ...range, message, severity }),
    setSource: (diagnostic, source) => { diagnostic.source = source; },
    setCode: (diagnostic, code) => { if (code != null) diagnostic.code = code; },
};

function diagnose(uri: string, source: string): FakeDiagnostic[] {
    const server = new LSPServer();
    server.didOpen(uri, source);
    const collection = new FakeDiagnosticCollection();
    publishDiagnostics(uri, server.getDiagnostics(uri), collection, fakeApi, (key, args) =>
        `${key}${args.length ? ` ${args.join(',')}` : ''}`);
    return collection.get(uri);
}

// The same source is exercised by the engine and by the diagnostic adapter.
{
    const source = String.raw`Option Explicit
Function Verify() As String
    Dim value As Long
    value = 2 + 3
    Verify = CStr(value)
End Function`;
    const engine = evalVBASingle(source);
    assert.strictEqual(engine.callProcedure('Verify', []), '5', 'engine fixture output');
    assert.strictEqual(diagnose('file:///workspace/engine.bas', source).length, 0,
        'valid engine fixture has no diagnostics');
    console.log('[PASS] Extension harness shares a valid engine fixture');
}

// The adapter preserves VS Code-facing metadata, including ranges and codes.
{
    const source = String.raw`Option Explicit
Sub Verify()
    missing = 1
End Sub`;
    const diagnostics = diagnose('file:///workspace/invalid.bas', source);
    assert.ok(diagnostics.some(d => d.message.includes("Variable 'missing' not declared")
        && d.source === 'vba-runner'), 'undeclared diagnostic metadata is preserved');
    console.log('[PASS] Extension harness preserves diagnostic metadata');
}

console.log('\n✅ Extension diagnostics harness: all tests passed');
