/** 財務関数の型変換・極小率マトリクス。 */
import { evalVBASingle, assert, assertClose } from '../../test-libs/test-runner';

const ev = evalVBASingle(String.raw`
Public pmtText, pmtInteger, pvTiny, fvTiny, rateTinyErr
Sub Test()
    pmtText = PMT("0.01", "12", "-1000", "0", "0")
    pmtInteger = PMT(CInt(0), CInt(4), CInt(100), CInt(-200), CInt(0))
    pvTiny = PV(0.000000000001, 4, -25, 100)
    fvTiny = FV(0.000000000001, 4, -25, -100)
    On Error Resume Next
    Dim ignored As Double
    ignored = Rate(CInt(4), CInt(-25), CInt(100), CInt(0), CInt(0), 0.000000000001)
    rateTinyErr = Err.Number
End Sub
`);
ev.callProcedure('Test', []);

assertClose(ev.env.get('pmttext') as number, 88.8487886783416, {
    message: 'PMT coerces numeric strings',
    // The host and runner may differ by a few ulps in the annuity factor.
    absolute: 1e-12,
});
assertClose(ev.env.get('pmtinteger') as number, 25, {
    message: 'PMT accepts Integer arguments',
});
assertClose(ev.env.get('pvtiny') as number, 0.008890058234065598, {
    message: 'PV remains stable at an extremely small rate',
});
assertClose(ev.env.get('fvtiny') as number, 200.00889005863414, {
    message: 'FV remains stable at an extremely small rate',
});
assert.strictEqual(ev.env.get('ratetinyerr'), 0,
    'Rate accepts a convergent solution with a tiny guess');
console.log('[PASS] Financial coercion and tiny-rate matrix');
