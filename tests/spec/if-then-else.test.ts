/**
 * If Statement (§5.4.2.8 & §5.4.2.9) のテスト
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

// --- 1. 基本的な If...Then...Else (複数行) ---
const basicIfCode = `
    Function CheckSign(n)
        If n > 0 Then
            CheckSign = "Positive"
        ElseIf n < 0 Then
            CheckSign = "Negative"
        Else
            CheckSign = "Zero"
        End If
    End Function
`;
assert.strictEqual(runFunc(basicIfCode, 'CheckSign', [10]), "Positive", '10 は Positive');
assert.strictEqual(runFunc(basicIfCode, 'CheckSign', [-5]), "Negative", '-5 は Negative');
assert.strictEqual(runFunc(basicIfCode, 'CheckSign', [0]), "Zero", '0 は Zero');
console.log('[PASS] If...Then...Else (複数行) 基本動作');

// --- 2. Single-line If ---
const singleLineIfCode = `
    Function FastAbs(n)
        FastAbs = n: If n < 0 Then FastAbs = -n
    End Function
`;
assert.strictEqual(runFunc(singleLineIfCode, 'FastAbs', [5]), 5, '5 は 5');
assert.strictEqual(runFunc(singleLineIfCode, 'FastAbs', [-10]), 10, '-10 は 10');
console.log('[PASS] Single-line If');

// --- 3. Single-line If...Then...Else ---
const singleLineElseCode = `
    Function Parrot(n)
        If n = 1 Then Parrot = "One" Else Parrot = "Other"
    End Function
`;
assert.strictEqual(runFunc(singleLineElseCode, 'Parrot', [1]), "One", '1 は One');
assert.strictEqual(runFunc(singleLineElseCode, 'Parrot', [2]), "Other", '2 は Other');
console.log('[PASS] Single-line If...Then...Else');

// --- 4. ネストした If ---
const nestedIfCode = `
    Function NestedIf(a, b)
        If a > 0 Then
            If b > 0 Then
                NestedIf = "Both Positive"
            Else
                NestedIf = "A Positive, B Not"
            End If
        Else
            NestedIf = "A Not Positive"
        End If
    End Function
`;
assert.strictEqual(runFunc(nestedIfCode, 'NestedIf', [1, 1]), "Both Positive", '1, 1');
assert.strictEqual(runFunc(nestedIfCode, 'NestedIf', [1, 0]), "A Positive, B Not", '1, 0');
assert.strictEqual(runFunc(nestedIfCode, 'NestedIf', [0, 1]), "A Not Positive", '0, 1');
console.log('[PASS] ネストした If');

console.log('\n✅ If...Then...Else: 全テスト通過');
