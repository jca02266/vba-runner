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

function runFunc(code: string, name: string, args: any[] = []): any {
    return evalVBA(code).callProcedure(name, args);
}

// 1. 中間引数を省略するとデフォルト値が使われる
{
    const r = runFunc(`
        Function Add3(Optional a As Long = 1, Optional b As Long = 2, Optional c As Long = 3)
            Add3 = a + b + c
        End Function
        Function Test()
            Test = Add3(10,,30)
        End Function
    `, 'Test');
    assert.strictEqual(r, 42, 'Add3(10,,30) = 10+2+30 = 42');
    console.log('[PASS] 中間省略: デフォルト値適用:', r);
}

// 2. 末尾引数を省略するとデフォルト値が使われる
{
    const r = runFunc(`
        Function Greet(name As String, Optional greeting As String = "Hello")
            Greet = greeting & " " & name
        End Function
        Function Test()
            Test = Greet("World",)
        End Function
    `, 'Test');
    assert.strictEqual(r, 'Hello World', 'Greet("World",) = "Hello World"');
    console.log('[PASS] 末尾省略: デフォルト値適用:', r);
}

// 3. 先頭引数を省略するとデフォルト値が使われる
{
    const r = runFunc(`
        Function Test()
            Test = Func(,20)
        End Function
        Function Func(Optional a As Long = 10, Optional b As Long = 0)
            Func = a + b
        End Function
    `, 'Test');
    assert.strictEqual(r, 30, 'Func(,20) = 10+20 = 30');
    console.log('[PASS] 先頭省略: デフォルト値適用:', r);
}

// 4. IsMissing は省略スロットで True になる（デフォルト値なし Optional）
{
    const r = runFunc(`
        Function Test()
            Test = Check(1,,3)
        End Function
        Function Check(Optional a, Optional b, Optional c)
            If IsMissing(b) Then
                Check = "missing"
            Else
                Check = "present"
            End If
        End Function
    `, 'Test');
    assert.strictEqual(r, 'missing', 'IsMissing(b) = True when ,, used');
    console.log('[PASS] 省略スロットで IsMissing = True:', r);
}

// 5. 省略なしの通常呼び出しは従来通り
{
    const r = runFunc(`
        Function Add3(Optional a As Long = 1, Optional b As Long = 2, Optional c As Long = 3)
            Add3 = a + b + c
        End Function
        Function Test()
            Test = Add3(10, 20, 30)
        End Function
    `, 'Test');
    assert.strictEqual(r, 60, 'Add3(10,20,30) = 60');
    console.log('[PASS] 省略なし通常呼び出し:', r);
}

// 6. キーワード引数と省略スロットの混在
{
    const r = runFunc(`
        Function Func(Optional a As Long = 1, Optional b As Long = 2, Optional c As Long = 3)
            Func = a * 100 + b * 10 + c
        End Function
        Function Test()
            Test = Func(,, c:=9)
        End Function
    `, 'Test');
    assert.strictEqual(r, 129, 'Func(,,c:=9) = 1*100+2*10+9 = 129');
    console.log('[PASS] 省略スロット + キーワード引数:', r);
}

console.log('\n✅ missing-arg: 全テスト通過');
