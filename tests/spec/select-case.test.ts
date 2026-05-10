/**
 * Select Case Statement のテスト (§5.4.2.10)
 *
 * 仕様書の range-clause 3種を網羅:
 *   1. expression        : Case 1, 2, 3
 *   2. start To end      : Case 1 To 5
 *   3. [Is] comparison   : Case Is > 10
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

// --- 1. 単純な値マッチ ---
const basicCode = `
Function GetLabel(x)
    Select Case x
        Case 1
            GetLabel = "one"
        Case 2
            GetLabel = "two"
        Case 3
            GetLabel = "three"
        Case Else
            GetLabel = "other"
    End Select
End Function
`;
assert.strictEqual(runFunc(basicCode, 'GetLabel', [1]), 'one',   'Case 1');
assert.strictEqual(runFunc(basicCode, 'GetLabel', [2]), 'two',   'Case 2');
assert.strictEqual(runFunc(basicCode, 'GetLabel', [3]), 'three', 'Case 3');
assert.strictEqual(runFunc(basicCode, 'GetLabel', [9]), 'other', 'Case Else');
console.log('[PASS] 単純な値マッチ');

// --- 2. カンマ区切り複数値 ---
const multiCode = `
Function IsWeekend(day)
    Select Case day
        Case 1, 7
            IsWeekend = True
        Case Else
            IsWeekend = False
    End Select
End Function
`;
assert.strictEqual(runFunc(multiCode, 'IsWeekend', [1]), true,  'day=1 is weekend');
assert.strictEqual(runFunc(multiCode, 'IsWeekend', [7]), true,  'day=7 is weekend');
assert.strictEqual(runFunc(multiCode, 'IsWeekend', [3]), false, 'day=3 is not weekend');
console.log('[PASS] カンマ区切り複数値');

// --- 3. To による範囲マッチ ---
const rangeCode = `
Function GetGrade(score)
    Select Case score
        Case 90 To 100
            GetGrade = "A"
        Case 70 To 89
            GetGrade = "B"
        Case 50 To 69
            GetGrade = "C"
        Case Else
            GetGrade = "F"
    End Select
End Function
`;
assert.strictEqual(runFunc(rangeCode, 'GetGrade', [95]),  'A', 'score 95 -> A');
assert.strictEqual(runFunc(rangeCode, 'GetGrade', [90]),  'A', 'score 90 -> A (境界)');
assert.strictEqual(runFunc(rangeCode, 'GetGrade', [100]), 'A', 'score 100 -> A (境界)');
assert.strictEqual(runFunc(rangeCode, 'GetGrade', [75]),  'B', 'score 75 -> B');
assert.strictEqual(runFunc(rangeCode, 'GetGrade', [55]),  'C', 'score 55 -> C');
assert.strictEqual(runFunc(rangeCode, 'GetGrade', [40]),  'F', 'score 40 -> F');
console.log('[PASS] To による範囲マッチ');

// --- 4. Is による比較マッチ ---
const isCode = `
Function Classify(n)
    Select Case n
        Case Is < 0
            Classify = "negative"
        Case Is = 0
            Classify = "zero"
        Case Is > 0
            Classify = "positive"
    End Select
End Function
`;
assert.strictEqual(runFunc(isCode, 'Classify', [-5]), 'negative', 'Is < 0');
assert.strictEqual(runFunc(isCode, 'Classify', [0]),  'zero',     'Is = 0');
assert.strictEqual(runFunc(isCode, 'Classify', [3]),  'positive', 'Is > 0');
console.log('[PASS] Is による比較マッチ');

// --- 5. 最初にマッチした Case のみ実行される（フォールスルーなし）---
const firstMatchCode = `
Function FirstMatch(x)
    Select Case x
        Case 1 To 10
            FirstMatch = "low"
        Case 5 To 20
            FirstMatch = "high"
    End Select
End Function
`;
assert.strictEqual(runFunc(firstMatchCode, 'FirstMatch', [5]), 'low', '最初のマッチのみ実行');
console.log('[PASS] 最初にマッチした Case のみ実行');

// --- 6. Case Else がない場合にマッチしなければ何もしない ---
const noElseCode = `
Function NoElse(x)
    NoElse = "default"
    Select Case x
        Case 1
            NoElse = "one"
    End Select
End Function
`;
assert.strictEqual(runFunc(noElseCode, 'NoElse', [1]), 'one',     'マッチあり');
assert.strictEqual(runFunc(noElseCode, 'NoElse', [2]), 'default', 'マッチなし、デフォルト値を維持');
console.log('[PASS] Case Else なし');

// --- 7. 文字列マッチ ---
const strCode = `
Function DayType(day)
    Select Case day
        Case "Mon", "Tue", "Wed", "Thu", "Fri"
            DayType = "weekday"
        Case "Sat", "Sun"
            DayType = "weekend"
        Case Else
            DayType = "unknown"
    End Select
End Function
`;
assert.strictEqual(runFunc(strCode, 'DayType', ['Mon']), 'weekday', 'Mon is weekday');
assert.strictEqual(runFunc(strCode, 'DayType', ['Sat']), 'weekend', 'Sat is weekend');
assert.strictEqual(runFunc(strCode, 'DayType', ['xyz']), 'unknown', 'unknown day');
console.log('[PASS] 文字列マッチ');

console.log('\n✅ Select Case: 全テスト通過');
