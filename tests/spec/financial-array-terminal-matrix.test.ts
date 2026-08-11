/** 財務関数の配列・償却終端・符号反転マトリクス。 */
import { evalVBASingle, assert, assertClose } from '../../test-libs/test-runner';

const ev = evalVBASingle(String.raw`
Public npvOneBased, npvZeroBased, irrAllPositiveErr
Public slnFlat, sydLast, ddbFlat, ipmtInverted, ppmtInverted
Sub Test()
    Dim oneBased(1 To 3) As Double, zeroBased(0 To 2) As Double
    oneBased(1) = -100: oneBased(2) = 40: oneBased(3) = 80
    zeroBased(0) = -100: zeroBased(1) = 40: zeroBased(2) = 80
    npvOneBased = NPV(0.1, oneBased)
    npvZeroBased = NPV(0.1, zeroBased)
    On Error Resume Next
    Dim positive(0 To 2) As Double
    positive(0) = 1: positive(1) = 2: positive(2) = 3
    Dim ignored As Double
    ignored = IRR(positive)
    irrAllPositiveErr = Err.Number
    Err.Clear
    slnFlat = SLN(100, 100, 5)
    sydLast = SYD(100, 0, 5, 5)
    ddbFlat = DDB(100, 100, 5, 1)
    ipmtInverted = IPmt(0.01, 2, 12, -1000)
    ppmtInverted = PPmt(0.01, 2, 12, -1000)
End Sub
`);
ev.callProcedure('Test', []);

assertClose(ev.env.get('npvonebased') as number, ev.env.get('npvzerobased') as number, {
    message: 'NPV preserves one-based array shape',
});
assert.strictEqual(ev.env.get('irrallpositiveerr'), 5,
    'IRR rejects a cash-flow array without a sign change');
assert.strictEqual(ev.env.get('slnflat'), 0, 'SLN returns zero when cost equals salvage');
assertClose(ev.env.get('sydlast') as number, 6.666666666666667, {
    message: 'SYD final period uses the terminal depreciation fraction',
});
assert.strictEqual(ev.env.get('ddbflat'), 0, 'DDB returns zero when cost equals salvage');
assert.ok((ev.env.get('ipmtinverted') as number) > 0,
    'IPmt reverses sign with a negative present value');
assert.ok((ev.env.get('ppmtinverted') as number) > 0,
    'PPmt reverses sign with a negative present value');
console.log('[PASS] Financial array, terminal, and sign matrix');
