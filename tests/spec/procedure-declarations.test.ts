import { Lexer } from '../../src/compiler/lexer';
import { Parser } from '../../src/compiler/parser';
import { Evaluator } from '../../src/compiler/evaluator';
import { assert } from '../ts/test-runner';

function runFunc(code: string, name: string, args: any[] = []): any {
    const tokens = new Lexer(code).tokenize();
    const ast = new Parser(tokens).parse();
    const ev = new Evaluator(console.log);
    ev.evaluate(ast);
    return ev.callProcedure(name, args);
}

console.log("--- Starting Procedure Declaration Tests ---");

// 1. Basic Function
assert.strictEqual(runFunc('Function Add(a, b) : Add = a + b : End Function', 'Add', [1, 2]), 3, 'Basic Function');

// 2. Function with Explicit Return Type (parsed, ignored at runtime)
assert.strictEqual(runFunc('Function GetStr() As String : GetStr = "Hello" : End Function', 'GetStr'), "Hello", 'Function with return type');

// 3. Basic Sub (returns vbaEmpty/null)
assert.strictEqual(runFunc('Sub MySub() : End Sub', 'MySub'), null, 'Basic Sub');

// 4. Sub with side effect
{
    const code = `
        Public x
        Sub SetX(v)
            x = v
        End Sub
    `;
    const tokens = new Lexer(code).tokenize();
    const ast = new Parser(tokens).parse();
    const ev = new Evaluator(console.log);
    ev.evaluate(ast);
    ev.callProcedure('SetX', [100]);
    assert.strictEqual(ev.env.get('x'), 100, 'Sub side effect');
}

// 5. Recursive Function
{
    const code = `
        Function Fact(n)
            If n <= 1 Then
                Fact = 1
            Else
                Fact = n * Fact(n - 1)
            End If
        End Function
    `;
    assert.strictEqual(runFunc(code, 'Fact', [5]), 120, 'Recursive Function');
}

console.log("✅ Procedure Declarations: All tests passed!");
