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

// 5. With ブロック内で JS prototype 上の getter/メソッドが解決できること
// 旧実装は Object.keys()（own プロパティのみ）を使っていたため、
// TypeScript class の get accessor（prototype 上に定義）を .Prop 構文で解決できなかった。
{
    // prototype に getter を持つ JS クラスを外部オブジェクトとして登録
    class MockFindLike {
        _text: string = '';
        _replacement = { Text: '' };

        get Text(): string { return this._text; }
        set Text(v: any) { this._text = String(v); }

        /** prototype getter — 旧実装では With ブロック内から .Replacement でアクセスできなかった */
        get Replacement(): { Text: string } { return this._replacement; }

        Execute() {}
    }

    const code = `
        Function TestWithPrototypeGetter() As String
            Dim obj As Object
            Set obj = CreateObject("MockFindLike")
            With obj
                .Text = "foo"
                .Replacement.Text = "bar"
            End With
            TestWithPrototypeGetter = obj.Text & "|" & obj.Replacement.Text
        End Function
    `;
    const ev = evalVBA(code);
    ev.registerExternalObject('MockFindLike', () => new MockFindLike());
    assert.strictEqual(
        ev.callProcedure('TestWithPrototypeGetter', []),
        'foo|bar',
        'With ブロック内で prototype getter (.Replacement) が解決できること'
    );
    console.log('[PASS] With block: prototype getter (.Replacement) resolution');
}

// 6. With ブロック内で prototype の no-arg メソッドが auto-call されること
{
    class MockAutoCall {
        _count: number = 0;
        /** no-arg メソッド — With ブロック内で .Count とアクセスすると auto-call される */
        Count() { return this._count; }
        Increment() { this._count++; }
    }

    const code = `
        Function TestWithAutoCall() As Long
            Dim obj As Object
            Set obj = CreateObject("MockAutoCall")
            obj.Increment
            obj.Increment
            With obj
                TestWithAutoCall = .Count
            End With
        End Function
    `;
    const ev = evalVBA(code);
    ev.registerExternalObject('MockAutoCall', () => new MockAutoCall());
    assert.strictEqual(
        ev.callProcedure('TestWithAutoCall', []),
        2,
        'With ブロック内で prototype の no-arg メソッドが auto-call されること'
    );
    console.log('[PASS] With block: prototype no-arg method auto-call');
}

console.log('\n✅ With Statement: 全テスト通過');
