import { Evaluator } from '../../src/engine/evaluator';
import { evalVBASingle, assert } from '../../test-libs/test-runner';

function evalVBA(code: string): Evaluator {
    return evalVBASingle(code);
}

const code = `
Sub TestByRef(ByRef x)
    x = x * 2
End Sub

Sub TestByVal(ByVal x)
    x = x * 2
End Sub

Sub TestDefaultIsByRef(x)
    x = x * 2
End Sub

Function TestByRefFunction(x)
    x = x + 10
    TestByRefFunction = x
End Function

Dim globalVar
Sub ModifyGlobal()
    globalVar = 10
    Call TestByRef(globalVar)
End Sub
`;

// Test 1: ByRef modifies caller variable
{
    const ev = evalVBA(code);
    ev.evalExpression(`
        Dim x
        x = 10
        Call TestByRef(x)
    `);
    assert.strictEqual(ev.evalExpression('x'), 20);
    console.log("✅ ByRef modifies caller variable");
}

// Test 2: ByVal does NOT modify caller variable
{
    const ev = evalVBA(code);
    ev.evalExpression(`
        Dim x
        x = 10
        Call TestByVal(x)
    `);
    assert.strictEqual(ev.evalExpression('x'), 10);
    console.log("✅ ByVal does not modify caller variable");
}

// Test 3: Default is ByRef
{
    const ev = evalVBA(code);
    ev.evalExpression(`
        Dim x
        x = 10
        Call TestDefaultIsByRef(x)
    `);
    assert.strictEqual(ev.evalExpression('x'), 20);
    console.log("✅ Default is ByRef");
}

// Test 4: Parentheses force ByVal evaluation
{
    const ev = evalVBA(code);
    ev.evalExpression(`
        Dim x
        x = 10
        Call TestByRef((x))
    `);
    assert.strictEqual(ev.evalExpression('x'), 10);
    console.log("✅ Parentheses force ByVal evaluation");
}

// Test 5: ByRef in function used in expressions
{
    const ev = evalVBA(code);
    ev.evalExpression(`
        Dim x
        x = 10
        Dim result
        result = TestByRefFunction(x) + x
    `);
    assert.strictEqual(ev.evalExpression('result'), 40);
    console.log("✅ ByRef in function used in expressions");
}

// Test 6: ByRef for module level variables
{
    const ev = evalVBA(code);
    ev.evalExpression('Call ModifyGlobal');
    assert.strictEqual(ev.evalExpression('globalVar'), 20);
    console.log("✅ ByRef for module level variables");
}

// Bug BX: ByVal UDT パラメーターがコピーを作らず参照として渡されていた
{
    const code = `
Type Point
    X As Long
    Y As Long
End Type
Sub MoveByRef(p As Point, dx As Long)
    p.X = p.X + dx
End Sub
Sub MoveByVal(ByVal p As Point, dx As Long)
    p.X = p.X + dx
End Sub
Function TestBX() As String
    Dim pt As Point
    pt.X = 10
    MoveByRef pt, 3
    Dim x1 As Long : x1 = pt.X  ' 13 (ByRef: caller's pt modified)
    MoveByVal pt, 10
    Dim x2 As Long : x2 = pt.X  ' 13 (ByVal: copy, caller's pt unchanged)
    TestBX = x1 & "," & x2
End Function
`;
    const result = evalVBA(code).callProcedure('TestBX', []);
    assert.strictEqual(result, '13,13', 'Bug BX: ByVal UDT creates a copy; caller is not affected');
    console.log('✅ Bug BX: ByVal UDT パラメーターのコピーセマンティクス');
}

console.log("--- All ByRef tests passed! ---");
