/**
 * VBARunner.run() / .eval() が VbaCurrency / VbaDecimal をオブジェクトのまま返すことを確認する
 * レグレッションテスト（Bug C-2 修正確認含む）
 *
 * Bug C-2: run() が Currency 型 ByRef パラメーターへの書き戻し後に
 *          JSON.stringify(BigInt) でクラッシュしていた。
 *          → formatVbaArg() で BigInt 含む型を JSON.stringify 回避
 *
 * 注: VbaCurrency / VbaDecimal は精度損失を避けるため number に変換せずオブジェクトのまま返す。
 *     呼び出し側は .toString() または .value で値を取り出す。
 */
import { VBARunner, assert } from '../../test-libs/test-runner';
import { VbaCurrency, VbaDecimal } from '../../src/engine/vba-types';

// --- Currency 戻り値は VbaCurrency オブジェクトのまま返る ---
{
    const r = new VBARunner([], { quiet: true });
    r.eval(`
Function GetCur() As Currency
    GetCur = CCur(0.3)
End Function
`);
    const result = r.run('GetCur', []);
    assert.ok(result instanceof VbaCurrency, 'Currency 戻り値は VbaCurrency インスタンス');
    assert.strictEqual(result.toString(), '0.3', 'VbaCurrency.toString() = "0.3"');
}
console.log('[PASS] Currency 戻り値は VbaCurrency オブジェクトのまま返る');

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
    assert.ok(result instanceof VbaCurrency, 'Bug C-2: 戻り値は VbaCurrency インスタンス');
    assert.strictEqual(result.toString(), '1100', 'Bug C-2: 1000 * 1.1 = 1100');
}
console.log('[PASS] Bug C-2: Currency ByRef 書き戻し後にクラッシュしない');

// --- Currency の精度: 0.1 + 0.2 = 0.3 (VBA 内で固定小数点演算) ---
{
    const r = new VBARunner([], { quiet: true });
    r.eval(`
Function SumCur(a As Currency, b As Currency) As Currency
    SumCur = a + b
End Function
`);
    const result = r.run('SumCur', [0.1, 0.2]);
    assert.ok(result instanceof VbaCurrency, 'Currency 和は VbaCurrency インスタンス');
    assert.strictEqual(result.toString(), '0.3', 'CCur(0.1) + CCur(0.2) = "0.3" (精度損失なし)');
}
console.log('[PASS] CCur(0.1) + CCur(0.2) の toString は "0.3" (固定小数点精度)');

// --- Decimal 戻り値は VbaDecimal オブジェクトのまま返る ---
{
    const r = new VBARunner([], { quiet: true });
    r.eval(`
Function GetDec() As Variant
    GetDec = CDec(1.5)
End Function
`);
    const result = r.run('GetDec', []);
    assert.ok(result instanceof VbaDecimal, 'Decimal 戻り値は VbaDecimal インスタンス');
    assert.strictEqual(result.toString(), '1.5', 'VbaDecimal.toString() = "1.5"');
}
console.log('[PASS] Decimal 戻り値は VbaDecimal オブジェクトのまま返る');

// --- Decimal の高精度: 28桁文字列が保持される ---
{
    const r = new VBARunner([], { quiet: true });
    r.eval(`
Function GetHiPrec() As Variant
    GetHiPrec = CDec("0.1000000000000000000000000001")
End Function
`);
    const result = r.run('GetHiPrec', []);
    assert.ok(result instanceof VbaDecimal, 'VbaDecimal インスタンス');
    assert.strictEqual(result.toString(), '0.1000000000000000000000000001', '28桁精度が保持される');
}
console.log('[PASS] Decimal 28桁精度: toString() が完全な文字列を返す');

// --- Decimal の高精度: CDec(1)/CDec(3) が 28桁 ---
{
    const r = new VBARunner([], { quiet: true });
    r.eval(`
Function DecimalDivision() As Variant
    DecimalDivision = CDec(1) / CDec(3)
End Function
`);
    const result = r.run('DecimalDivision', []);
    assert.ok(result instanceof VbaDecimal, 'VbaDecimal インスタンス');
    assert.ok(result.toString().startsWith('0.3333333333333333333'), 'CDec(1)/CDec(3) が 19桁以上の精度');
}
console.log('[PASS] Decimal 精度: CDec(1)/CDec(3) が高精度 VbaDecimal を返す');

console.log('\n✅ VbaCurrency/VbaDecimal オブジェクト保持 (Bug C-2): 全テスト通過');
