/**
 * VBARunner.run() / .eval() が VbaCurrency / VbaDecimal を JS number に正規化することを確認する
 * レグレッションテスト（Bug C-1 / C-2）
 *
 * Bug C-1: run() / eval() が VbaCurrency を VbaCurrency オブジェクトのまま返し、
 *          JS number との比較が失敗していた。
 * Bug C-2: run() が Currency 型 ByRef パラメーターへの書き戻し後に
 *          JSON.stringify(BigInt) でクラッシュしていた。
 */
import { VBARunner, assert } from '../../test-libs/test-runner';

// --- Bug C-1: Currency 戻り値が JS number に変換される ---
{
    const r = new VBARunner([], { quiet: true });
    r.eval(`
Function GetCur() As Currency
    GetCur = CCur(0.3)
End Function
`);
    const result = r.run('GetCur', []);
    assert.strictEqual(typeof result, 'number', 'Bug C-1: Currency 戻り値の typeof は number');
    assert.strictEqual(result, 0.3, 'Bug C-1: CCur(0.3) の戻り値は JS number 0.3');
}
console.log('[PASS] Bug C-1: Currency 戻り値が JS number に変換される');

// --- Bug C-2: Currency 型パラメーター ByRef 書き戻し後にクラッシュしない ---
{
    const r = new VBARunner([], { quiet: true });
    r.eval(`
Function CalcWithTax(price As Currency) As Currency
    CalcWithTax = price * CCur(1.1)
End Function
`);
    let threw = false;
    let result: any;
    try {
        result = r.run('CalcWithTax', [1000]);
    } catch {
        threw = true;
    }
    assert.ok(!threw, 'Bug C-2: Currency ByRef 書き戻しでクラッシュしない');
    assert.strictEqual(typeof result, 'number', 'Bug C-2: 戻り値の typeof は number');
    assert.strictEqual(result, 1100, 'Bug C-2: 1000 * 1.1 = 1100');
}
console.log('[PASS] Bug C-2: Currency ByRef 書き戻し後にクラッシュしない');

// --- eval() も VbaCurrency を number に変換する ---
{
    const r = new VBARunner([], { quiet: true });
    r.eval(`
Function SumCur(a As Currency, b As Currency) As Currency
    SumCur = a + b
End Function
`);
    const result = r.run('SumCur', [0.1, 0.2]);
    assert.strictEqual(typeof result, 'number', 'Currency 和の typeof は number');
    assert.strictEqual(result, 0.3, 'CCur(0.1) + CCur(0.2) = 0.3 (厳密一致)');
}
console.log('[PASS] CCur(0.1) + CCur(0.2) = 0.3 (JS number として厳密一致)');

// --- Decimal 戻り値も JS number に変換される ---
{
    const r = new VBARunner([], { quiet: true });
    r.eval(`
Function GetDec() As Variant
    GetDec = CDec(1.5)
End Function
`);
    const result = r.run('GetDec', []);
    assert.strictEqual(typeof result, 'number', 'Decimal 戻り値の typeof は number');
    assert.strictEqual(result, 1.5, 'CDec(1.5) の戻り値は JS number 1.5');
}
console.log('[PASS] Decimal 戻り値が JS number に変換される');

// --- Currency の精度: 0.1 + 0.2 = 0.3 (VBA 内で CStr して検証) ---
{
    const r = new VBARunner([], { quiet: true });
    r.eval(`
Function CurrencyPrecision() As String
    Dim a As Currency, b As Currency
    a = CCur(0.1)
    b = CCur(0.2)
    CurrencyPrecision = CStr(a + b)
End Function
`);
    const result = r.run('CurrencyPrecision', []);
    assert.strictEqual(result, '0.3', 'Currency: CStr(CCur(0.1) + CCur(0.2)) = "0.3"');
}
console.log('[PASS] Currency 精度: CCur(0.1) + CCur(0.2) の CStr は "0.3"');

// --- Decimal の精度: 1/3 が 28桁の文字列になるか ---
{
    const r = new VBARunner([], { quiet: true });
    r.eval(`
Function DecimalDivision() As String
    Dim d As Variant
    d = CDec(1) / CDec(3)
    DecimalDivision = CStr(d)
End Function
`);
    const result = r.run('DecimalDivision', []);
    assert.ok(result.startsWith('0.3333333333333333333'), 'Decimal: CDec(1)/CDec(3) が 19桁以上の精度');
}
console.log('[PASS] Decimal 精度: CDec(1)/CDec(3) が高精度文字列を返す');

console.log('\n✅ VbaCurrency/VbaDecimal 正規化 (Bug C-1/C-2): 全テスト通過');
