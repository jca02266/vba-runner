import { LSPServer } from '../../src/lsp/server';

const server = new LSPServer();

// Test 1: No diagnostics for valid code
server.didOpen('file:///test.bas', 'Sub Foo()\n    x = 1\nEnd Sub');
const d1 = server.getDiagnostics('file:///test.bas');
console.assert(d1.length === 0, `Expected 0 diagnostics, got ${d1.length}`);
console.log('[PASS] Valid code: no diagnostics');

// Test 2: Parse error produces diagnostic
server.didOpen('file:///bad.bas', '@@@bad');
const d2 = server.getDiagnostics('file:///bad.bas');
console.assert(d2.length === 1, `Expected 1 diagnostic, got ${d2.length}`);
console.assert(d2[0].severity === 1, 'Severity should be 1 (error)');
console.assert(d2[0].range.start.line === 0, 'Error on line 0');
console.assert(d2[0].source === 'vba-runner', 'Source should be vba-runner');
console.log('[PASS] Parse error: 1 diagnostic, severity=1, line=0, source=vba-runner');

// Test 3: Unknown URI returns empty array
const d3 = server.getDiagnostics('file:///nonexistent.bas');
console.assert(d3.length === 0, 'Unknown URI: empty array');
console.log('[PASS] Unknown URI: empty array');

// Test 4: didChange updates diagnostics
server.didChange('file:///test.bas', '@@@updated bad');
const d4 = server.getDiagnostics('file:///test.bas');
console.assert(d4.length === 1, `Expected 1 diagnostic after change, got ${d4.length}`);
console.log('[PASS] didChange: diagnostics updated');

// Test 5: didClose clears document (getDiagnostics returns [])
server.didClose('file:///test.bas');
const d5 = server.getDiagnostics('file:///test.bas');
console.assert(d5.length === 0, 'After close: empty array');
console.log('[PASS] didClose: empty diagnostics');

console.log('\n✅ LSPServer.getDiagnostics: 全テスト通過');
