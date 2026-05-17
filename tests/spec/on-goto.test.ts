import { Lexer } from '../../src/engine/lexer';
import { Parser } from '../../src/engine/parser';
import { Evaluator } from '../../src/engine/evaluator';
import { assert } from '../../test-libs/test-runner';

function runFunc(code: string, name: string, args: any[] = []): any {
    const tokens = new Lexer(code).tokenize();
    const ast = new Parser(tokens).parse();
    const ev = new Evaluator(() => {});
    ev.evaluate(ast);
    return ev.callProcedure(name, args);
}

// 1. On...GoTo Statement
{
    const code = `
        Function TestOnGoTo(idx)
            On idx GoTo Label1, Label2, Label3
            TestOnGoTo = 0
            Exit Function
        Label1:
            TestOnGoTo = 1
            Exit Function
        Label2:
            TestOnGoTo = 2
            Exit Function
        Label3:
            TestOnGoTo = 3
        End Function
    `;
    assert.strictEqual(runFunc(code, 'TestOnGoTo', [1]), 1, 'On 1 GoTo Label1');
    assert.strictEqual(runFunc(code, 'TestOnGoTo', [2]), 2, 'On 2 GoTo Label2');
    assert.strictEqual(runFunc(code, 'TestOnGoTo', [3]), 3, 'On 3 GoTo Label3');
    assert.strictEqual(runFunc(code, 'TestOnGoTo', [0]), 0, 'On 0 GoTo should skip');
    assert.strictEqual(runFunc(code, 'TestOnGoTo', [4]), 0, 'On 4 GoTo should skip');
}

// 2. GoSub and Return
{
    const code = `
        Function TestGoSub()
            Dim x
            x = 10
            GoSub MySub
            x = x + 1
            TestGoSub = x
            Exit Function
        MySub:
            x = x * 2
            Return
        End Function
    `;
    assert.strictEqual(runFunc(code, 'TestGoSub'), 21, 'GoSub should jump and Return should come back');
}

// 3. Nested GoSub
{
    const code = `
        Function TestNestedGoSub()
            Dim x
            x = 1
            GoSub Sub1
            TestNestedGoSub = x
            Exit Function
        Sub1:
            x = x + 10
            GoSub Sub2
            Return
        Sub2:
            x = x * 2
            Return
        End Function
    `;
    assert.strictEqual(runFunc(code, 'TestNestedGoSub'), 22, 'Nested GoSub should work with LIFO stack');
}

// 4. On...GoSub
{
    const code = `
        Function TestOnGoSub(idx)
            Dim x
            x = 100
            On idx GoSub S1, S2
            TestOnGoSub = x
            Exit Function
        S1:
            x = x + 1
            Return
        S2:
            x = x + 2
            Return
        End Function
    `;
    assert.strictEqual(runFunc(code, 'TestOnGoSub', [1]), 101, 'On 1 GoSub S1');
    assert.strictEqual(runFunc(code, 'TestOnGoSub', [2]), 102, 'On 2 GoSub S2');
}

console.log('\n✅ On...GoTo/GoSub & GoSub/Return: 全テスト通過');
