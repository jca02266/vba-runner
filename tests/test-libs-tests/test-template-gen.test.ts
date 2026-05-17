import { Lexer } from '../../src/engine/lexer';
import { Parser } from '../../src/engine/parser';
import { findIoSideEffects, vbaTypeToTs } from '../../test-libs/vba-analyzer';
import { assert } from '../../test-libs/test-runner';

function parseProc(code: string): any {
    const tokens = new Lexer(code).tokenize();
    const ast = new Parser(tokens).parse();
    return ast.body.find((s: any) => s.type === 'ProcedureDeclaration');
}

// ---- vbaTypeToTs ----

{
    assert.strictEqual(vbaTypeToTs('Long'),    'number', 'Long -> number');
    assert.strictEqual(vbaTypeToTs('Integer'), 'number', 'Integer -> number');
    assert.strictEqual(vbaTypeToTs('Double'),  'number', 'Double -> number');
    assert.strictEqual(vbaTypeToTs('Single'),  'number', 'Single -> number');
    assert.strictEqual(vbaTypeToTs('Currency'),'number', 'Currency -> number');
    assert.strictEqual(vbaTypeToTs('Byte'),    'number', 'Byte -> number');
    assert.strictEqual(vbaTypeToTs('String'),  'string', 'String -> string');
    assert.strictEqual(vbaTypeToTs('Boolean'), 'boolean','Boolean -> boolean');
    assert.strictEqual(vbaTypeToTs('Variant'), 'any',    'Variant -> any');
    assert.strictEqual(vbaTypeToTs('Object'),  'any',    'Object -> any');
    assert.strictEqual(vbaTypeToTs(null),      'any',    'null -> any');
    console.log('[PASS] vbaTypeToTs');
}

// ---- returnType in AST ----

{
    const proc = parseProc(`
Function CalcTotal(a As Long, b As Long) As Long
    CalcTotal = a + b
End Function
`);
    assert.strictEqual(proc.returnType, 'Long', 'returnType を保存する');
    console.log('[PASS] parser: returnType を AST に保存');
}

{
    const proc = parseProc(`
Function GetName() As String
    GetName = "hello"
End Function
`);
    assert.strictEqual(proc.returnType, 'String', 'String 戻り値型');
    console.log('[PASS] parser: String 戻り値型');
}

{
    const proc = parseProc(`
Sub DoWork()
    Dim x As Long
End Sub
`);
    assert.strictEqual(proc.returnType, undefined, 'Sub は returnType なし');
    console.log('[PASS] parser: Sub は returnType なし');
}

{
    const proc = parseProc(`
Function Calc() As Boolean
    Calc = True
End Function
`);
    assert.strictEqual(proc.returnType, 'Boolean', 'Boolean 戻り値型');
    console.log('[PASS] parser: Boolean 戻り値型');
}

// ---- findIoSideEffects ----

{
    const proc = parseProc(`
Sub ShowMessage()
    MsgBox "hello"
End Sub
`);
    assert.strictEqual(findIoSideEffects(proc.body), 1, 'MsgBox を検出');
    console.log('[PASS] findIoSideEffects: MsgBox');
}

{
    const proc = parseProc(`
Sub AskUser()
    Dim v As String
    v = InputBox("Enter value")
End Sub
`);
    assert.strictEqual(findIoSideEffects(proc.body), 1, 'InputBox を検出');
    console.log('[PASS] findIoSideEffects: InputBox');
}

{
    const proc = parseProc(`
Sub LogValue(x As Long)
    Debug.Print x
End Sub
`);
    assert.strictEqual(findIoSideEffects(proc.body), 1, 'Debug.Print を検出');
    console.log('[PASS] findIoSideEffects: Debug.Print');
}

{
    const proc = parseProc(`
Sub MultiIo()
    MsgBox "a"
    Debug.Print "b"
    MsgBox "c"
End Sub
`);
    assert.strictEqual(findIoSideEffects(proc.body), 3, '複数の I/O 呼び出しを集計');
    console.log('[PASS] findIoSideEffects: 複数件');
}

{
    const proc = parseProc(`
Function PureCalc(a As Long) As Long
    PureCalc = a * 2
End Function
`);
    assert.strictEqual(findIoSideEffects(proc.body), 0, 'I/O なしは 0');
    console.log('[PASS] findIoSideEffects: I/O なし');
}

console.log('\n✅ test-template-gen: 全テスト通過');
