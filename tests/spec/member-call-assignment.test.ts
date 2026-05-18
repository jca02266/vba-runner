import { Lexer } from '../../src/engine/lexer';
import { Parser } from '../../src/engine/parser';
import { Evaluator } from '../../src/engine/evaluator';
import { assert } from '../../test-libs/test-runner';

function evalVBA(code: string): any {
    const tokens = new Lexer(code).tokenize();
    const ast = new Parser(tokens).parse();
    const ev = new Evaluator(console.log);
    ev.evaluate(ast);
    return ev;
}

function runFunc(code: string, name: string, args: any[] = []): any {
    return evalVBA(code).callProcedure(name, args);
}

// Test 1: dict.Item(key) = val（バグ再現: evaluateAssignmentToVariable が callee=MemberExpression を拒否していた）
{
    const result = runFunc(`
Function F() As String
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Item("x") = "hello"
    F = d.Item("x")
End Function
`, 'F');
    assert.strictEqual(result, 'hello', 'dict.Item(key) = val: Property Let via member access');
    console.log('[PASS] dict.Item(key) = val');
}

// Test 2: 複数キーの連続代入
{
    const result = runFunc(`
Function F() As Long
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Item("a") = 1
    d.Item("b") = 2
    d.Item("c") = 3
    F = d.Item("a") + d.Item("b") + d.Item("c")
End Function
`, 'F');
    assert.strictEqual(result, 6, 'dict.Item multiple keys sum = 6');
    console.log('[PASS] dict.Item multiple keys');
}

// Test 3: ループ内で dict.Item(key) = val
{
    const result = runFunc(`
Function F() As Long
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    Dim i As Long
    For i = 1 To 5
        d.Item(CStr(i)) = i * 10
    Next i
    F = d.Item("3")
End Function
`, 'F');
    assert.strictEqual(result, 30, 'dict.Item(key) = val in loop: key "3" = 30');
    console.log('[PASS] dict.Item in loop');
}

// Test 4: 上書き代入
{
    const result = runFunc(`
Function F() As String
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Item("k") = "first"
    d.Item("k") = "second"
    F = d.Item("k")
End Function
`, 'F');
    assert.strictEqual(result, 'second', 'dict.Item(key) = val overwrite');
    console.log('[PASS] dict.Item overwrite');
}

console.log('\n✅ member-call-assignment: 全テスト通過');
