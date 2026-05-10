import { Lexer } from '../../src/compiler/lexer';
import { Parser } from '../../src/compiler/parser';
import { Evaluator } from '../../src/compiler/evaluator';
import { assert } from '../ts/test-runner';

function makeEvaluator(code: string): Evaluator {
    const tokens = new Lexer(code).tokenize();
    const ast = new Parser(tokens).parse();
    const ev = new Evaluator(console.log);
    ev.evaluate(ast);
    return ev;
}

function runFunc(ev: Evaluator, name: string, args: any[] = []): any {
    return ev.callProcedure(name, args);
}

// --- 1. Static変数（非staticプロシージャ内）---
{
    const code = `
    Function Counter()
        Static n As Integer
        n = n + 1
        Counter = n
    End Function
    `;
    const ev = makeEvaluator(code);
    assert.strictEqual(runFunc(ev, 'Counter'), 1, '1回目: n=1');
    assert.strictEqual(runFunc(ev, 'Counter'), 2, '2回目: n=2');
    assert.strictEqual(runFunc(ev, 'Counter'), 3, '3回目: n=3');
    console.log('[PASS] Static変数（非staticプロシージャ）');
}

// --- 2. Static Sub（すべての変数が永続）---
{
    const code = `
    Static Sub Accumulate()
        Dim total As Long
        total = total + 10
    End Sub
    Function GetTotal()
        Static total As Long
        total = total + 10
        GetTotal = total
    End Function
    `;
    const ev = makeEvaluator(code);
    assert.strictEqual(runFunc(ev, 'GetTotal'), 10, '1回目: total=10');
    assert.strictEqual(runFunc(ev, 'GetTotal'), 20, '2回目: total=20');
    assert.strictEqual(runFunc(ev, 'GetTotal'), 30, '3回目: total=30');
    console.log('[PASS] Static Function（すべての変数が永続）');
}

// --- 3. Static Functionで複数の静的変数 ---
{
    const code = `
    Function MultiStatic()
        Static a As Integer
        Static b As Integer
        a = a + 1
        b = b + 2
        MultiStatic = a * 10 + b
    End Function
    `;
    const ev = makeEvaluator(code);
    assert.strictEqual(runFunc(ev, 'MultiStatic'), 12,  '1回目: a=1,b=2');
    assert.strictEqual(runFunc(ev, 'MultiStatic'), 24,  '2回目: a=2,b=4');
    assert.strictEqual(runFunc(ev, 'MultiStatic'), 36,  '3回目: a=3,b=6');
    console.log('[PASS] 複数のStatic変数');
}

// --- 4. Static SubとFunction（initial-static形式）---
{
    const code = `
    Static Function InitialStatic()
        Dim x As Integer
        x = x + 5
        InitialStatic = x
    End Function
    `;
    const ev = makeEvaluator(code);
    assert.strictEqual(runFunc(ev, 'InitialStatic'), 5,  '1回目: x=5');
    assert.strictEqual(runFunc(ev, 'InitialStatic'), 10, '2回目: x=10');
    console.log('[PASS] Static Function（initial-static形式）');
}

// --- 5. 異なるプロシージャ間でStatic変数は独立 ---
{
    const code = `
    Function FuncA()
        Static n As Integer
        n = n + 1
        FuncA = n
    End Function
    Function FuncB()
        Static n As Integer
        n = n + 10
        FuncB = n
    End Function
    `;
    const ev = makeEvaluator(code);
    assert.strictEqual(runFunc(ev, 'FuncA'), 1,  'FuncA 1回目');
    assert.strictEqual(runFunc(ev, 'FuncB'), 10, 'FuncB 1回目');
    assert.strictEqual(runFunc(ev, 'FuncA'), 2,  'FuncA 2回目（FuncBとは独立）');
    assert.strictEqual(runFunc(ev, 'FuncB'), 20, 'FuncB 2回目（FuncAとは独立）');
    console.log('[PASS] 異なるプロシージャ間でStatic変数は独立');
}

console.log('\n✅ Static キーワード: 全テスト通過');
