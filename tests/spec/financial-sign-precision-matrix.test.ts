/** 財務関数の符号・ゼロ率・高精度境界マトリクス。 */
import { evalVBASingle, assert, assertClose } from '../../test-libs/test-runner';

const ev = evalVBASingle(String.raw`
Public fvZero, pvZero, pmtZero, fvPos, fvNeg, pvPos, pvNeg
Public currencyPmt, currencyPv
Sub Test()
    fvZero = FV(0, 4, -25, -100)
    pvZero = PV(0, 4, -25, 100)
    pmtZero = PMT(0, 4, 100, -200)
    fvPos = FV(0.01, 12, -100, -1000)
    fvNeg = FV(0.01, 12, 100, 1000)
    pvPos = PV(0.01, 12, -100, -1000)
    pvNeg = PV(0.01, 12, 100, 1000)
    currencyPmt = PMT(CCur("0.0001"), 120, CCur("-123456789012.3456"))
    currencyPv = PV(CCur("0.0001"), 120, CCur("-123456789.0123"), CCur("456789012.3456"))
End Sub
`);
ev.callProcedure('Test', []);

assertClose(ev.env.get('fvzero') as number, 200, {
    message: 'FV zero-rate sign convention',
});
assertClose(ev.env.get('pvzero') as number, 0, {
    message: 'PV zero-rate cash-flow cancellation',
});
assertClose(ev.env.get('pmtzero') as number, 25, {
    message: 'PMT zero-rate sign convention',
});

const fvPos = ev.env.get('fvpos') as number;
const fvNeg = ev.env.get('fvneg') as number;
const pvPos = ev.env.get('pvpos') as number;
const pvNeg = ev.env.get('pvneg') as number;
assertClose(fvPos, -fvNeg, { message: 'FV is odd under cash-flow sign inversion' });
assertClose(pvPos, -pvNeg, { message: 'PV is odd under cash-flow sign inversion' });

assert.ok(Number.isFinite(ev.env.get('currencypmt') as number),
    'PMT remains finite for Currency-sized principal');
assert.ok(Number.isFinite(ev.env.get('currencypv') as number),
    'PV remains finite for Currency-sized cash flows');
console.log('[PASS] Financial sign, zero-rate, and precision matrix');
