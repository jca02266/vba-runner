/**
 * For Statement (§5.4.2.3) & Exit For (§5.4.2.5) のテスト
 */
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

// --- 1. 基本的な For ループ ---
const basicForCode = `
    Function SumFor(n)
        Dim s
        Dim i
        s = 0
        For i = 1 To n
            s = s + i
        Next i
        SumFor = s
    End Function
`;
assert.strictEqual(runFunc(basicForCode, 'SumFor', [5]), 15, '1+2+3+4+5 = 15');
assert.strictEqual(runFunc(basicForCode, 'SumFor', [0]), 0, '1 To 0 は実行されない');
console.log('[PASS] For Statement 基本動作');

// --- 2. Step 指定の For ループ ---
const stepForCode = `
    Function SumStep(n)
        Dim s
        Dim i
        s = 0
        For i = 1 To n Step 2
            s = s + i
        Next
        SumStep = s
    End Function
`;
assert.strictEqual(runFunc(stepForCode, 'SumStep', [5]), 9, '1+3+5 = 9');
assert.strictEqual(runFunc(stepForCode, 'SumStep', [6]), 9, '1+3+5 = 9 (6は超える)');
console.log('[PASS] For Statement Step 対応');

// --- 3. マイナスの Step ---
const backStepCode = `
    Function Backwards(n)
        Dim s
        Dim i
        s = 0
        For i = n To 1 Step -1
            s = s + i
        Next
        Backwards = s
    End Function
`;
assert.strictEqual(runFunc(backStepCode, 'Backwards', [5]), 15, '5+4+3+2+1 = 15');
console.log('[PASS] For Statement マイナス Step');

// --- 4. Exit For ---
const exitForCode = `
    Function FindFirstEven(n)
        Dim i
        FindFirstEven = -1
        For i = 1 To n
            If i Mod 2 = 0 Then
                FindFirstEven = i
                Exit For
            End If
        Next
    End Function
`;
assert.strictEqual(runFunc(exitForCode, 'FindFirstEven', [10]), 2, '2でループを抜ける');
console.log('[PASS] Exit For');

// --- 5. ネストした For と Exit For ---
const nestedForCode = `
    Function NestedExit()
        Dim count
        Dim i, j
        count = 0
        For i = 1 To 3
            For j = 1 To 10
                count = count + 1
                If j = 5 Then Exit For
            Next
        Next
        NestedExit = count
    End Function
`;
assert.strictEqual(runFunc(nestedForCode, 'NestedExit'), 15, '内側ループを3回途中で抜ける (3 * 5 = 15)');
console.log('[PASS] ネストした For と Exit For');

console.log('\n✅ For...Next & Exit For: 全テスト通過');
