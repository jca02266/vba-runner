import { LSPServer } from '../../src/lsp/server';

const server = new LSPServer();

// Test 1: No diagnostics for valid code
server.didOpen('file:///test.bas', 'Option Explicit\nSub Foo()\n    Dim x As Long\n    x = 1\n    Debug.Print x\nEnd Sub');
const d1 = server.getDiagnostics('file:///test.bas');
console.assert(d1.length === 0, `Expected 0 diagnostics, got ${d1.length}`);
console.log('[PASS] Valid code: no diagnostics');

// Test 2: Parse error produces diagnostic
server.didOpen('file:///bad.bas', '@@@bad');
const d2 = server.getDiagnostics('file:///bad.bas');
const parseErrors2 = d2.filter((diagnostic: any) => diagnostic.source === 'vba-runner');
console.assert(parseErrors2.length === 1, `Expected 1 parse diagnostic, got ${parseErrors2.length}`);
console.assert(parseErrors2[0].severity === 1, 'Severity should be 1 (error)');
console.assert(parseErrors2[0].range.start.line === 0, 'Error on line 0');
console.assert(parseErrors2[0].source === 'vba-runner', 'Source should be vba-runner');
console.log('[PASS] Parse error: 1 diagnostic, severity=1, line=0, source=vba-runner');

// Test 3: Unknown URI returns empty array
const d3 = server.getDiagnostics('file:///nonexistent.bas');
console.assert(d3.length === 0, 'Unknown URI: empty array');
console.log('[PASS] Unknown URI: empty array');

// Test 4: didChange updates diagnostics
server.didChange('file:///test.bas', '@@@updated bad');
const d4 = server.getDiagnostics('file:///test.bas');
const parseErrors4 = d4.filter((diagnostic: any) => diagnostic.source === 'vba-runner');
console.assert(parseErrors4.length === 1, `Expected 1 parse diagnostic after change, got ${parseErrors4.length}`);
console.log('[PASS] didChange: diagnostics updated');

// Test 5: didClose releases the open document while workspace diagnostics remain
server.didClose('file:///test.bas');
const d5 = server.getDiagnostics('file:///test.bas');
const parseErrors5 = d5.filter((diagnostic: any) => diagnostic.source === 'vba-runner');
console.assert(parseErrors5.length === 1, 'After close: workspace diagnostics remain available');
console.log('[PASS] didClose: workspace diagnostics remain available');

// Test 6: Range 変数の添字アクセスに Hint が出る
server.didOpen('file:///range.bas', [
    'Sub Test()',
    '    Dim rng As Range',
    '    rng(1, 1) = 99',
    'End Sub',
].join('\n'));
const d6 = server.getDiagnostics('file:///range.bas');
const rangeHints = d6.filter((d: any) => d.source === 'vba-dataflow');
console.assert(rangeHints.length >= 1, `Expected >= 1 Range hint, got ${rangeHints.length}`);
console.assert(rangeHints[0].severity === 4, 'Severity should be 4 (Hint)');
console.assert(rangeHints[0].message.includes('rng'), `Message should mention 'rng': ${rangeHints[0].message}`);
console.log('[PASS] Range 添字アクセスに Hint が出る');

// Test 7: Range 変数がない場合は vba-dataflow Hint が出ない
server.didOpen('file:///norange.bas', [
    'Sub Test()',
    '    Dim n As Long',
    '    n = 42',
    'End Sub',
].join('\n'));
const d7 = server.getDiagnostics('file:///norange.bas');
const noRangeHints = d7.filter((d: any) => d.source === 'vba-dataflow');
console.assert(noRangeHints.length === 0, `Expected 0 Range hints, got ${noRangeHints.length}`);
console.log('[PASS] 非 Range コードには vba-dataflow Hint が出ない');

// Test 8: Property Let の暗黙 ByRef 値引数は VS Code Warning になる
server.didOpen('file:///property-let.bas', [
    'Property Let Value(index As Long, value As String)',
    'End Property',
].join('\n'));
const d8 = server.getDiagnostics('file:///property-let.bas');
const propertyWarnings = d8.filter((d: any) => d.code === 'VBA015');
console.assert(propertyWarnings.length === 1, `Expected 1 VBA015 warning, got ${propertyWarnings.length}`);
console.assert(propertyWarnings[0].severity === 2, 'VBA015 severity should be Warning (2)');
console.log('[PASS] Property Let 暗黙 ByRef: VS Code Warning');

// Test 9: VBE-exported member attributes are valid in class diagnostics
server.didOpen('file:///ExcelQueueDefaultValue.cls', [
    'VERSION 1.0 CLASS',
    'BEGIN',
    '  MultiUse = -1',
    'END',
    'Attribute VB_Name = "ExcelQueueDefaultValue"',
    'Option Explicit',
    'Public Property Get Value() As Variant',
    'Attribute Value.VB_UserMemId = 0',
    '    Value = 1',
    'End Property',
].join('\n'));
const classAttributeDiagnostics = server.getDiagnostics('file:///ExcelQueueDefaultValue.cls')
    .filter((d: any) => d.source === 'vba-runner');
console.assert(classAttributeDiagnostics.length === 0,
    `Expected no class attribute diagnostics, got ${classAttributeDiagnostics.map((d: any) => d.message).join('; ')}`);
console.log('[PASS] VBE class member Attribute is accepted by diagnostics');

// Test 10: Cross-module member qualifiers are known to Option Explicit
server.loadWorkspaceFile('file:///workspace/ExcelQueueQualifiedProperty.cls', String.raw`VERSION 1.0 CLASS
Attribute VB_Name = "ExcelQueueQualifiedProperty"
Option Explicit
Public Property Get Name() As String
    Name = "module-a"
End Property
Public Property Get Values() As Variant
    Values = Array(10, 20, 30)
End Property`);
server.didOpen('file:///workspace/qualified-property.bas', String.raw`Option Explicit
Sub Verify()
    Dim value As Variant, errNo As Long
    Err.Clear: value = ExcelQueueQualifiedProperty.Name: errNo = Err.Number
    Err.Clear: value = ExcelQueueQualifiedProperty.Values(1): errNo = Err.Number
End Sub`);
const qualifiedPropertyDiagnostics = server.getDiagnostics('file:///workspace/qualified-property.bas')
    .filter((d: any) => d.source === 'vba-runner');
console.assert(qualifiedPropertyDiagnostics.length === 0,
    `Expected no cross-module qualifier diagnostics, got ${qualifiedPropertyDiagnostics.map((d: any) => d.message).join('; ')}`);

server.didOpen('file:///workspace/unknown-qualifier.bas', String.raw`Option Explicit
Sub Verify()
    Dim value As Variant
    value = UnknownQualifiedModule.Name
End Sub`);
const unknownQualifierDiagnostics = server.getDiagnostics('file:///workspace/unknown-qualifier.bas')
    .filter((d: any) => d.source === 'vba-runner');
console.assert(unknownQualifierDiagnostics.some((d: any) => d.message.includes("'UnknownQualifiedModule'")),
    'Unknown module qualifiers must still be reported');
console.log('[PASS] Cross-module qualifiers are known while unknown qualifiers remain diagnostics');

// Test 11: Host globals recommend a mock and provide an AI-oriented guide link
server.didOpen('file:///workspace/host-dependent.bas', String.raw`Option Explicit
Sub VerifyHost()
    Dim value As Variant
    value = ThisWorkbook.Name
    value = Application.ActiveSheet.Name
    value = CurrentDb()
End Sub`);
const hostMockDiagnostics = server.getDiagnostics('file:///workspace/host-dependent.bas')
    .filter((d: any) => d.source === 'vba-mock-advisor');
console.assert(hostMockDiagnostics.length === 3,
    `Expected 3 host mock recommendations, got ${hostMockDiagnostics.length}`);
console.assert(hostMockDiagnostics.every((d: any) => d.message.includes('__mocks__')
    && d.message.includes('docs/guides/MOCK_GUIDE.md')
    && d.message.includes('AI prompt example')
    && d.message.includes('full current VBA source')
    && d.message.includes('Do not guess')),
    'Host recommendations should require source context and include an AI prompt example');
console.log('[PASS] Host globals recommend __mocks__ and include an AI prompt example');

console.log('\n✅ LSPServer.getDiagnostics: 全テスト通過');
