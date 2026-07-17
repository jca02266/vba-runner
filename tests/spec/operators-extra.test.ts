/**
 * Operators (Extra) (§5.6.9) のテスト
 */
import { evalVBASingle, assert } from '../../test-libs/test-runner';

function evalVBA(code: string): any {
    return evalVBASingle(code);
}

// --- 1. 文字列連結 (&) ---
const concatCode = `
    Public s
    Sub Test()
        s = "Hello" & " " & 123 & True
    End Sub
`;
const ev1 = evalVBA(concatCode);
ev1.callProcedure('Test', []);
assert.strictEqual(ev1.env.get('s'), "Hello 123True", '文字列連結 (&)');
console.log('[PASS] 文字列連結 (&)');

// --- 2. 比較演算子 (=, <>, <, >, <=, >=) ---
const compCode = `
    Public eq, ne, lt, gt, le, ge
    Sub Test()
        eq = (10 = 10)
        ne = (10 <> 20)
        lt = (10 < 20)
        gt = (20 > 10)
        le = (10 <= 10)
        ge = (10 >= 10)
    End Sub
`;
const ev2 = evalVBA(compCode);
ev2.callProcedure('Test', []);
assert.strictEqual(ev2.env.get('eq').value, -1, '10 = 10');
assert.strictEqual(ev2.env.get('ne').value, -1, '10 <> 20');
assert.strictEqual(ev2.env.get('lt').value, -1, '10 < 20');
assert.strictEqual(ev2.env.get('gt').value, -1, '20 > 10');
assert.strictEqual(ev2.env.get('le').value, -1, '10 <= 10');
assert.strictEqual(ev2.env.get('ge').value, -1, '10 >= 10');
console.log('[PASS] 比較演算子');

// --- 3. Like 演算子 ---
const likeCode = `
    Public res1, res2, res3, res4
    Sub Test()
        res1 = ("abcde" Like "a*e")
        res2 = ("abcde" Like "a?c?e")
        res3 = ("abcde" Like "a#cde") ' Fail: # は数字 [0-9]
        res4 = ("abcde" Like "[a-z]bcde")
    End Sub
`;
const ev3 = evalVBA(likeCode);
ev3.callProcedure('Test', []);
assert.strictEqual(ev3.env.get('res1').value, -1, '"abcde" Like "a*e"');
assert.strictEqual(ev3.env.get('res2').value, -1, '"abcde" Like "a?c?e"');
assert.strictEqual(ev3.env.get('res3').value, 0, '"abcde" Like "a#cde" (False)');
assert.strictEqual(ev3.env.get('res4').value, -1, '"abcde" Like "[a-z]bcde"');
console.log('[PASS] Like 演算子');

// --- Bug BH: `\` と Mod が浮動小数点の被演算子を事前丸めしていなかった ---
{
    const intDivCode = `
        Public res1, res2, res3, res4, res5, res6
        Sub Test()
            res1 = 7.5 \\ 2        ' 8 \\ 2 = 4 (7.5 → 銀行家丸め→8)
            res2 = 6.5 \\ 2        ' 6 \\ 2 = 3 (6.5 → 銀行家丸め→6)
            res3 = -7.5 \\ 2       ' -8 \\ 2 = -4
            res4 = 7.5 Mod 2       ' 8 Mod 2 = 0
            res5 = 6.5 Mod 3       ' 6 Mod 3 = 0
            res6 = 7.3 \\ 2        ' 7 \\ 2 = 3 (7.3 → 7)
        End Sub
    `;
    const evIntDiv = evalVBA(intDivCode);
    evIntDiv.callProcedure('Test', []);
    assert.strictEqual(evIntDiv.env.get('res1'), 4, '7.5 \\ 2 = 4 (銀行家丸め後)');
    assert.strictEqual(evIntDiv.env.get('res2'), 3, '6.5 \\ 2 = 3 (銀行家丸め後)');
    assert.strictEqual(evIntDiv.env.get('res3'), -4, '-7.5 \\ 2 = -4');
    assert.strictEqual(evIntDiv.env.get('res4'), 0, '7.5 Mod 2 = 0 (8 Mod 2)');
    assert.strictEqual(evIntDiv.env.get('res5'), 0, '6.5 Mod 3 = 0 (6 Mod 3)');
    assert.strictEqual(evIntDiv.env.get('res6'), 3, '7.3 \\ 2 = 3 (7.3→7, 7\\2)');
    console.log('[PASS] Bug BH: \\ と Mod の浮動小数点事前丸め');
}

// --- Bug BI: `2 ^ -1` が parse error になる（VBA では 0.5）---
{
    const powerCode = `
        Public res1, res2, res3
        Sub Test()
            res1 = 2 ^ -1
            res2 = 4 ^ -0.5
            res3 = 2 ^ -2
        End Sub
    `;
    const evPow = evalVBA(powerCode);
    evPow.callProcedure('Test', []);
    assert.strictEqual(evPow.env.get('res1'), 0.5, '2 ^ -1 = 0.5');
    assert.strictEqual(evPow.env.get('res2'), 0.5, '4 ^ -0.5 = 0.5');
    assert.strictEqual(evPow.env.get('res3'), 0.25, '2 ^ -2 = 0.25');
    console.log('[PASS] Bug BI: ^ 演算子の負の指数');
}

// --- Bug BJ: Currency混在の \\ / Mod が事前丸めと0除算チェックを欠いていた ---
{
    const curCode = `
        Public res1, res2, res3, res4, res5
        Sub Test()
            Dim c As Currency
            c = CCur(7.5)
            res1 = c \\ 2        ' 8 \\ 2 = 4 (7.5 → 銀行家丸め→8)
            res2 = c \\ 2.1      ' 8 \\ 2 = 4 (7.5→8, 2.1→2)
            c = CCur(6.5)
            res3 = c Mod 3       ' 6 Mod 3 = 0 (6.5→6)
            res4 = c Mod 3.1     ' 6 Mod 3 = 0 (6.5→6, 3.1→3)
            res5 = CCur(5.5) \\ CCur(3.5)  ' round(5.5)=6, round(3.5)=4, 6\\4=1
        End Sub
    `;
    const evCur = evalVBA(curCode);
    evCur.callProcedure('Test', []);
    assert.strictEqual(evCur.env.get('res1'), 4, 'CCur(7.5) \\ 2 = 4');
    assert.strictEqual(evCur.env.get('res2'), 4, 'CCur(7.5) \\ 2.1 = 4');
    assert.strictEqual(evCur.env.get('res3'), 0, 'CCur(6.5) Mod 3 = 0');
    assert.strictEqual(evCur.env.get('res4'), 0, 'CCur(6.5) Mod 3.1 = 0');
    assert.strictEqual(evCur.env.get('res5'), 1, 'CCur(5.5) \\ CCur(3.5) = 1');
    console.log('[PASS] Bug BJ: Currency混在の \\ / Mod 事前丸め');
}

// --- Bug BU: 文字列と数値の混在比較 "7" = 7 が False になっていた ---
{
    const code = `
Function Test() As String
    Dim r1 As Boolean, r2 As Boolean, r3 As Boolean, r4 As Boolean
    r1 = ("7" = 7)
    r2 = ("7" <> 7)
    r3 = (7 = "7")
    r4 = (7 <> "7")
    Test = r1 & "," & r2 & "," & r3 & "," & r4
End Function
`;
    const result = evalVBASingle(code).callProcedure('Test', []);
    assert.strictEqual(result, 'True,False,True,False', 'Bug BU: 文字列と数値の混在比較 = と <> が VBA 自動型変換を行う');
    console.log('[PASS] Bug BU: 文字列/数値混在の = と <> 比較');
}

// --- Bug 29-B: ^ は左結合 (2^3^2 = 64、右結合では 512) ---
{
    const ev = evalVBASingle('');
    assert.strictEqual(ev.evalExpression('2 ^ 3 ^ 2'), 64, 'Bug 29-B: 2^3^2 = (2^3)^2 = 64 (左結合)');
    assert.strictEqual(ev.evalExpression('2 ^ 3'), 8, 'Bug 29-B: 2^3 = 8');
    console.log('[PASS] Bug 29-B: ^ 演算子は左結合');
}

// --- Bug 29-C: 虚数乗算 (-8)^0.5 は Error 5 ---
{
    let errNum = 0;
    try { evalVBASingle('').evalExpression('(-8) ^ 0.5'); } catch (e: any) { errNum = e?.number ?? -1; }
    assert.strictEqual(errNum, 5, 'Bug 29-C: (-8)^0.5 → Error 5 (NaN は Invalid procedure call)');
    console.log('[PASS] Bug 29-C: (-8)^0.5 → Error 5');
}

// --- Bug 29-D: Null & Null は Null を返す ---
{
    const code = `
Function TestNullConcat()
    Dim r1, r2, r3, r4
    r1 = IsNull(Null & Null)
    r2 = (Null & "")
    r3 = ("" & Null)
    r4 = (Null & "a")
    TestNullConcat = r1 & "," & r2 & "," & r3 & "," & r4
End Function
`;
    const result = evalVBASingle(code).callProcedure('TestNullConcat', []);
    assert.strictEqual(result, 'True,,,a', 'Bug 29-D: Null & Null → Null; Null & "" → ""; "" & Null → ""; Null & "a" → "a"');
    console.log('[PASS] Bug 29-D: Null & Null → Null');
}

console.log('\n✅ Operators (Extra): 全テスト通過');
