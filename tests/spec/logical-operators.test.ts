/**
 * 論理/ビット演算子のテスト (§5.6.9.8)
 *
 * 対象:
 *   - And / Or / Xor / Eqv / Imp / Not (§5.6.9.8.1〜6)
 *   - Like 演算子 (§5.6.9.6)
 *   - 演算子の優先順位
 */
import { evalVBASingle, assert } from '../../test-libs/test-runner';

function evalVBA(code: string): any {
    return evalVBASingle(code);
}

function ev(expr: string): any {
    return evalVBA('').evalExpression(expr);
}

console.log('--- Starting Logical/Bitwise Operators Tests ---');

// --- And ---
{
    assert.strictEqual(ev('-1 And -1'), -1, 'True And True = True');
    assert.strictEqual(ev('-1 And 0'), 0, 'True And False = False');
    assert.strictEqual(ev('5 And 3'), 1, '101 And 011 = 001');
    console.log('[PASS] And');
}

// --- Or ---
{
    assert.strictEqual(ev('-1 Or 0'), -1, 'True Or False = True');
    assert.strictEqual(ev('0 Or 0'), 0, 'False Or False = False');
    assert.strictEqual(ev('5 Or 3'), 7, '101 Or 011 = 111');
    console.log('[PASS] Or');
}

// --- Xor ---
{
    assert.strictEqual(ev('-1 Xor -1'), 0, 'True Xor True = False');
    assert.strictEqual(ev('-1 Xor 0'), -1, 'True Xor False = True');
    assert.strictEqual(ev('5 Xor 3'), 6, '101 Xor 011 = 110');
    console.log('[PASS] Xor');
}

// --- Eqv ---
{
    assert.strictEqual(ev('-1 Eqv -1'), -1, 'True Eqv True = True');
    assert.strictEqual(ev('-1 Eqv 0'), 0, 'True Eqv False = False');
    assert.strictEqual(ev('5 Eqv 3'), -7, 'Not(101 Xor 011) = Not(110)');
    console.log('[PASS] Eqv');
}

// --- Imp ---
{
    assert.strictEqual(ev('-1 Imp -1'), -1, 'True Imp True = True');
    assert.strictEqual(ev('-1 Imp 0'), 0, 'True Imp False = False');
    assert.strictEqual(ev('0 Imp 0'), -1, 'False Imp False = True');
    assert.strictEqual(ev('5 Imp 3'), -5, '(Not 5) Or 3 = -6 Or 3 = -5');
    console.log('[PASS] Imp');
}

// --- Not ---
{
    assert.strictEqual(ev('Not -1'), 0, 'Not True = False');
    assert.strictEqual(ev('Not 0'), -1, 'Not False = True');
    assert.strictEqual(ev('Not 5'), -6, 'Not 5 = -6 (bitwise)');
    console.log('[PASS] Not');
}

// --- 演算子の優先順位: And > Or > Xor > Eqv > Imp ---
{
    assert.strictEqual(ev('-1 Or -1 And 0'), -1, 'And が Or より先 → -1 Or (−1 And 0) = -1 Or 0 = -1');
    assert.strictEqual(ev('0 Xor -1 Or -1'), -1, 'Or が Xor より先 → 0 Xor (-1 Or -1) = 0 Xor -1 = -1');
    assert.strictEqual(ev('-1 Eqv -1 Xor -1'), 0, 'Xor が Eqv より先 → -1 Eqv (-1 Xor -1) = -1 Eqv 0 = 0');
    assert.strictEqual(ev('0 Imp -1 Eqv 0'), -1, 'Eqv が Imp より先 → 0 Imp (-1 Eqv 0) = 0 Imp 0 = -1');
    console.log('[PASS] 演算子の優先順位');
}

// --- Like 演算子 ---
{
    assert.strictEqual(ev(`"abc" Like "a*"`), -1, '* ワイルドカード');
    assert.strictEqual(ev(`"abc" Like "?b?"`), -1, '? 任意の1文字');
    assert.strictEqual(ev(`"a1b" Like "a#b"`), -1, '# 任意の1桁数字');
    assert.strictEqual(ev(`"a2b" Like "a[123]b"`), -1, '[charlist] 文字クラス');
    assert.strictEqual(ev(`"a4b" Like "a[123]b"`), 0, '[charlist] にない場合は False');
    assert.strictEqual(ev(`"a4b" Like "a[!123]b"`), -1, '[!charlist] 否定');
    assert.strictEqual(ev(`"ABC" Like "abc"`), 0, 'デフォルトは Binary 比較で大文字小文字を区別');
    console.log('[PASS] Like');
}

// --- And / Not の字句境界 ---
{
    assert.strictEqual(ev('1 And 0'), 0, '空白付きAnd');
    assert.strictEqual(ev('1And 0'), 0, '左辺数値に隣接したAnd');
    assert.throws(() => ev('1 And0'), /Run-time error '35'|not defined/i,
        '式評価では右辺数値に隣接したAnd0が未定義識別子になる');
    assert.strictEqual(String(ev('Not 1 = 1')), 'False', 'Notは比較結果に適用');
    const procedureEval = evalVBASingle(String.raw`
Function ProbeAnd() As Long
    ProbeAnd = 1 And 0
End Function
Function ProbeAdjacent() As Long
    ProbeAdjacent = 1And 0
End Function
Function ProbeNot() As String
    ProbeNot = CStr(Not 1 = 1)
End Function
`);
    assert.strictEqual(procedureEval.callProcedure('ProbeAnd', []), 0,
        '手続き内の空白付きAnd');
    assert.strictEqual(procedureEval.callProcedure('ProbeAdjacent', []), 0,
        '手続き内の左辺隣接And');
    assert.strictEqual(procedureEval.callProcedure('ProbeNot', []), 'False',
        '手続き内のNot比較');
    console.log('[PASS] And / Not の字句境界');
}

console.log('\n✅ 論理/ビット演算子: 全テスト通過');
