import { Lexer } from '../../src/compiler/lexer';
import { Parser } from '../../src/compiler/parser';
import { Evaluator } from '../../src/compiler/evaluator';
import { assert } from '../ts/test-runner';

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

// 1. 基本的なWithステートメント (オブジェクト)
{
    const code = `
        Function TestWith()
            Dim obj
            Set obj = CreateObject("Scripting.Dictionary")
            With obj
                .Add "key1", "value1"
                .Add "key2", "value2"
            End With
            TestWith = obj.Exists("key1")
        End Function
    `;
    assert.isTrue(runFunc(code, 'TestWith'), 'With block should call methods on the object');
    console.log('[PASS] Basic With (Object)');
}

// 2. 入れ子のWithステートメント
{
    const code = `
        Function TestNestedWith()
            Dim obj1, obj2
            Set obj1 = CreateObject("Scripting.Dictionary")
            Set obj2 = CreateObject("Scripting.Dictionary")
            With obj1
                .Add "name", "obj1"
                With obj2
                    .Add "name", "obj2"
                End With
            End With
            TestNestedWith = obj1.Exists("name") And obj2.Exists("name")
        End Function
    `;
    assert.isTrue(runFunc(code, 'TestNestedWith'), 'Nested With blocks should work correctly');
    console.log('[PASS] Nested With');
}

// 3. ユーザ定義型(UDT)に対するWith
{
    const code = `
        Type Person
            Name As String
            Age As Integer
        End Type

        Function TestWithUDT()
            Dim p As Person
            With p
                .Name = "Claude"
                .Age = 3
            End With
            TestWithUDT = p.Name
        End Function
    `;
    assert.strictEqual(runFunc(code, 'TestWithUDT'), "Claude", 'With block should work with UDT');
    console.log('[PASS] With with UDT');
}

// 4. 式の結果に対するWith
{
    const code = `
        Function GetDict()
            Set GetDict = CreateObject("Scripting.Dictionary")
        End Function

        Function TestWithExpr()
            With GetDict()
                .Add "key", "value"
                TestWithExpr = .Exists("key")
            End With
        End Function
    `;
    assert.isTrue(runFunc(code, 'TestWithExpr'), 'With block should work with expressions');
    console.log('[PASS] With with Expression');
}

console.log('\n✅ With Statement: 全テスト通過');
