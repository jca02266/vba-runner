import { Lexer } from '../../src/engine/lexer';
import { Parser } from '../../src/engine/parser';
import { Evaluator } from '../../src/engine/evaluator';
import { assert } from '../../test-libs/test-runner';

function evalVBA(code: string): Evaluator {
    const tokens = new Lexer(code).tokenize();
    const ast = new Parser(tokens).parse();
    const ev = new Evaluator(console.log);
    ev.evaluate(ast);
    return ev;
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

console.log("--- All ByRef tests passed! ---");
