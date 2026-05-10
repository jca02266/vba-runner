/**
 * While Statement のテスト (§5.4.2.2)
 *
 * 網羅する観点:
 *   1. 基本動作（条件が真の間繰り返す）
 *   2. 条件が最初から偽（0回実行）
 *   3. ネストした While...Wend
 *   4. While 内の変数更新
 */
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

// --- 1. 基本動作: 1 から n まで加算 ---
const basicCode = `
Function SumWhile(n)
    Dim total
    Dim i
    total = 0
    i = 1
    While i <= n
        total = total + i
        i = i + 1
    Wend
    SumWhile = total
End Function
`;
assert.strictEqual(runFunc(basicCode, 'SumWhile', [5]), 15, '1+2+3+4+5 = 15');
assert.strictEqual(runFunc(basicCode, 'SumWhile', [1]), 1,  'n=1 のとき 1');
assert.strictEqual(runFunc(basicCode, 'SumWhile', [10]), 55, '1..10 の合計');
console.log('[PASS] 基本動作');

// --- 2. 条件が最初から偽（ループ本体は一度も実行されない）---
const zeroIterCode = `
Function NeverRuns()
    Dim x
    x = 0
    While x > 0
        x = x + 1
    Wend
    NeverRuns = x
End Function
`;
assert.strictEqual(runFunc(zeroIterCode, 'NeverRuns'), 0, '最初から偽なら本体未実行');
console.log('[PASS] 条件が最初から偽');

// --- 3. ネストした While...Wend ---
const nestedCode = `
Function CountPairs(n)
    Dim count
    Dim i
    Dim j
    count = 0
    i = 1
    While i <= n
        j = 1
        While j <= n
            count = count + 1
            j = j + 1
        Wend
        i = i + 1
    Wend
    CountPairs = count
End Function
`;
assert.strictEqual(runFunc(nestedCode, 'CountPairs', [3]), 9,  '3x3 = 9');
assert.strictEqual(runFunc(nestedCode, 'CountPairs', [4]), 16, '4x4 = 16');
console.log('[PASS] ネストした While...Wend');

// --- 4. 文字列の繰り返し構築 ---
const strCode = `
Function BuildStr(n)
    Dim result
    Dim i
    result = ""
    i = 0
    While i < n
        result = result & "x"
        i = i + 1
    Wend
    BuildStr = result
End Function
`;
assert.strictEqual(runFunc(strCode, 'BuildStr', [4]), 'xxxx', '4文字の繰り返し');
assert.strictEqual(runFunc(strCode, 'BuildStr', [0]), '',     '0回は空文字');
console.log('[PASS] 文字列の繰り返し構築');

console.log('\n✅ While Statement: 全テスト通過');
