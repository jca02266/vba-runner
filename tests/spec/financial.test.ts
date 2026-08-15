/**
 * Financial Functions (§6.1.2.6) のテスト
 */
import { evalVBASingle, assert, assertClose } from '../../test-libs/test-runner';

function evalVBA(code: string): any {
    return evalVBASingle(code);
}

// --- 1. FV, PV, PMT, NPer, Rate ---
const financialCode = `
    Public resFV, resPV, resPMT, resNPer, resRate
    Sub Test()
        ' Rate=5%, NPer=10, Pmt=-100, PV=-1000
        resFV = FV(0.05 / 12, 10 * 12, -100, -1000)
        resPV = PV(0.05 / 12, 10 * 12, -100, 20000)
        resPMT = Pmt(0.05 / 12, 10 * 12, -10000)
        resNPer = NPer(0.05 / 12, -100, -10000, 20000)
        resRate = Rate(10 * 12, -100, -10000, 20000)
    End Sub
`;
const ev1 = evalVBA(financialCode);
ev1.callProcedure('Test', []);
console.log('resFV:', ev1.env.get('resfv'));
console.log('resPV:', ev1.env.get('respv'));
console.log('resPMT:', ev1.env.get('respmt'));
assert.ok(ev1.env.get('resfv') > 0, 'FV');
assert.ok(Math.abs(ev1.env.get('respv') - (-2752)) < 100, 'PV');
assert.ok(Math.abs(ev1.env.get('respmt') - 106) < 10, 'PMT');
console.log('[PASS] 基本財務関数 (FV, PV, PMT, etc.)');

// Rate must accept an exact zero-rate solution even when Guess is explicitly 0.
{
    const ev = evalVBASingle(`
    Function ProbeZeroRate() As String
        On Error Resume Next
        Dim value As Double
        value = Rate(10, -10, 100, 0, 0, 0)
        ProbeZeroRate = CStr(Err.Number) & "|" & CStr(value)
    End Function
    `);
    assert.strictEqual(ev.callProcedure('ProbeZeroRate', []), '0|0',
        'Rate accepts an exact zero-rate cash-flow solution');
}
console.log('[PASS] Rate zero-rate solution');

// Payment Type is restricted to 0 (end) or 1 (beginning).
for (const expression of [
    'FV(0.01, 12, -100, 0, 2)',
    'PV(0.01, 12, -100, 0, 2)',
    'PMT(0.01, 12, 100, 0, 2)',
    'NPer(0.01, -100, 100, 0, 2)',
    'Rate(12, -100, 100, 0, 2)',
    'IPmt(0.01, 1, 12, 100, 0, 2)',
    'PPmt(0.01, 1, 12, 100, 0, 2)',
    'FV(0.01, 12, -100, 0, 0.5)',
]) {
    const ev = evalVBASingle(`Function Probe() As Long\n On Error Resume Next\n Dim value As Double\n value = ${expression}\n Probe = Err.Number\nEnd Function`);
    assert.strictEqual(ev.callProcedure('Probe', []), 5, `${expression} rejects invalid Type`);
}
console.log('[PASS] Financial Type enumeration validation');

// --- 2. SLN, SYD, DDB ---
const depCode = `
    Public resSLN, resSYD, resDDB
    Sub Test()
        ' Cost=10000, Salvage=1000, Life=5, Period=1
        resSLN = SLN(10000, 1000, 5)
        resSYD = SYD(10000, 1000, 5, 1)
        resDDB = DDB(10000, 1000, 5, 1)
    End Sub
`;
const ev2 = evalVBA(depCode);
ev2.callProcedure('Test', []);
assert.strictEqual(ev2.env.get('ressln'), 1800, 'SLN');
assert.strictEqual(ev2.env.get('ressyd'), 3000, 'SYD');
assert.strictEqual(ev2.env.get('resddb'), 4000, 'DDB');
console.log('[PASS] 減価償却関数 (SLN, SYD, DDB)');

// --- 3. IRR, MIRR, NPV ---
const cashCode = `
    Public resIRR, resNPV
    Sub Test()
        Dim v
        v = Array(-10000, 3000, 4200, 6800)
        resIRR = IRR(v)
        resNPV = NPV(0.1, v)
    End Sub
`;
const ev3 = evalVBA(cashCode);
ev3.callProcedure('Test', []);
assert.ok(Math.abs(ev3.env.get('resirr') - 0.16) < 0.05, 'IRR');
assert.ok(Math.abs(ev3.env.get('resnpv') - 1188) < 100, 'NPV');
console.log('[PASS] キャッシュフロー関数 (IRR, NPV)');

// Excel may retain a machine-epsilon residue where the runner rounds MIRR to
// zero. Compare this boundary with the shared absolute/relative policy rather
// than treating harmless floating-point operation order as a regression.
{
    const ev = evalVBASingle(String.raw`
    Public result
    Sub Test()
        Dim flows(0 To 1) As Double
        flows(0) = -100
        flows(1) = 100
        result = MIRR(flows, -1.1, 0.1)
    End Sub
    `);
    ev.callProcedure('Test', []);
    assertClose(ev.env.get('result') as number, 2.22044604925031e-16, {
        message: 'MIRR Excel epsilon boundary (XL-037)',
    });
}
console.log('[PASS] MIRR Excel epsilon boundary comparison');

// MIRR's exact -1 reinvestment boundary is an invalid argument (Excel Error 5).
{
    const ev = evalVBASingle(String.raw`
    Function ProbeMirrReinvestBoundary() As Long
        On Error Resume Next
        Dim ignored As Double
        ignored = MIRR(Array(-100, 100), 0.1, -1)
        ProbeMirrReinvestBoundary = Err.Number
    End Function
    `);
    assert.strictEqual(ev.callProcedure('ProbeMirrReinvestBoundary', []), 5,
        'MIRR ReinvestRate=-1 is Error 5');
}
console.log('[PASS] MIRR ReinvestRate boundary error');

// --- Bug 24-1: NPV が 1-based 配列で NaN を返す ---
// Dim flows(1 To N) で宣言した配列を渡すと vbaBase=1 のため
// values.map(Number) がインデックス 0 (undefined → NaN) を含んでいた
{
    const ev = evalVBASingle(`
    Public resNpv1, resNpv0
    Sub Test()
        Dim flows1(1 To 3) As Double
        flows1(1) = -30000 : flows1(2) = 40000 : flows1(3) = 50000
        resNpv1 = NPV(0.1, flows1)

        Dim flows0(2) As Double
        flows0(0) = -30000 : flows0(1) = 40000 : flows0(2) = 50000
        resNpv0 = NPV(0.1, flows0)
    End Sub
    `);
    ev.callProcedure('Test', []);
    const npv1 = ev.env.get('resnpv1') as number;
    const npv0 = ev.env.get('resnpv0') as number;
    assert.ok(!isNaN(npv1), `Bug 24-1: NPV 1-based should not be NaN (got ${npv1})`);
    assert.ok(Math.abs(npv1 - npv0) < 0.01, `NPV 1-based should equal 0-based (${npv1} vs ${npv0})`);
    assert.ok(Math.abs(npv1 - 43350.86) < 1, `NPV(0.1,[-30000,40000,50000]) ≈ 43350.86 (got ${npv1.toFixed(2)})`);
}
console.log('[PASS] Bug 24-1: NPV 1-based 配列で正常動作');

// IRR/MIRR must also ignore the unused slot before a 1-based VBA array.
{
    const ev = evalVBASingle(`
    Public irr1, irr0, mirr1, mirr0
    Sub Test()
        Dim flows1(1 To 4) As Double
        flows1(1) = -10000 : flows1(2) = 3000
        flows1(3) = 4200 : flows1(4) = 6800
        irr1 = IRR(flows1)
        mirr1 = MIRR(flows1, 0.1, 0.12)

        Dim flows0(3) As Double
        flows0(0) = -10000 : flows0(1) = 3000
        flows0(2) = 4200 : flows0(3) = 6800
        irr0 = IRR(flows0)
        mirr0 = MIRR(flows0, 0.1, 0.12)
    End Sub
    `);
    ev.callProcedure('Test', []);
    const irr1 = ev.env.get('irr1') as number;
    const irr0 = ev.env.get('irr0') as number;
    const mirr1 = ev.env.get('mirr1') as number;
    const mirr0 = ev.env.get('mirr0') as number;
    assert.ok(Number.isFinite(irr1), `IRR 1-based should be finite (got ${irr1})`);
    assert.ok(Number.isFinite(mirr1), `MIRR 1-based should be finite (got ${mirr1})`);
    assert.ok(Math.abs(irr1 - irr0) < 1e-10, `IRR 1-based should equal 0-based (${irr1} vs ${irr0})`);
    assert.ok(Math.abs(mirr1 - mirr0) < 1e-10, `MIRR 1-based should equal 0-based (${mirr1} vs ${mirr0})`);
}
console.log('[PASS] IRR/MIRR 1-based 配列で正常動作');

// IPmt/PPmt retain their signs after period 1 and validate the requested period.
{
    const ev = evalVBASingle(`
    Public ordinaryInterest, ordinaryPrincipal
    Public dueInterest, duePrincipal, outOfRangeError
    Sub Test()
        ordinaryInterest = IPmt(0.01, 12, 12, 1000)
        ordinaryPrincipal = PPmt(0.01, 12, 12, 1000)
        dueInterest = IPmt(0.01, 2, 12, 1000, 0, 1)
        duePrincipal = PPmt(0.01, 2, 12, 1000, 0, 1)
        On Error GoTo badPeriod
        ordinaryInterest = IPmt(0.01, 13, 12, 1000)
        Exit Sub
    badPeriod:
        outOfRangeError = Err.Number
    End Sub
    `);
    ev.callProcedure('Test', []);
    const closeTo = (actual: number, expected: number, message: string) =>
        assert.ok(Math.abs(actual - expected) < 1e-6, `${message} (${actual} vs ${expected})`);
    closeTo(ev.env.get('ordinaryInterest'), -0.8796909770132879, 'IPmt period 12');
    closeTo(ev.env.get('ordinaryPrincipal'), -87.96909770132839, 'PPmt period 12');
    closeTo(ev.env.get('dueInterest'), -9.120308677978415, 'IPmt due period 2');
    closeTo(ev.env.get('duePrincipal'), -78.84878902334997, 'PPmt due period 2');
    assert.strictEqual(ev.env.get('outOfRangeError'), 5, 'IPmt period > nper is Error 5');
}
console.log('[PASS] IPmt/PPmt period and payment timing');

// Bug 145-A: Pmt の NPer=0 が JavaScript の Infinity を返していた
{
    const ev = evalVBASingle(`
    Public pmtError
    Sub Test()
        On Error GoTo badNPer
        Dim ignored As Double
        ignored = Pmt(0.01, 0, 1000)
        Exit Sub
    badNPer:
        pmtError = Err.Number
    End Sub
    `);
    ev.callProcedure('Test', []);
    assert.strictEqual(ev.env.get('pmterror'), 5, 'Pmt NPer=0 is Error 5, not Infinity');
}
console.log('[PASS] Bug 145-A: Pmt のゼロ期間をエラー化');

// Bug 180-A: IRR/MIRR must reject cash flows without both signs.
{
    const ev = evalVBASingle(`
    Public irrError, mirrError
    Sub Test()
        Dim flows(0 To 2) As Double
        flows(0) = 100 : flows(1) = 20 : flows(2) = 10
        Dim ignored As Double
        On Error Resume Next
        ignored = IRR(flows)
        irrError = Err.Number
        Err.Clear
        ignored = MIRR(flows, 0.1, 0.1)
        mirrError = Err.Number
    End Sub
    `);
    ev.callProcedure('Test', []);
    assert.strictEqual(ev.env.get('irrerror'), 5, 'IRR without positive and negative flows is Error 5');
    assert.strictEqual(ev.env.get('mirrerror'), 5, 'MIRR without positive and negative flows is Error 5');
}
console.log('[PASS] Bug 180-A: IRR/MIRRの符号不足をエラー化');

// NPV's VBA contract requires at least one payment and one receipt, including
// arrays whose lower bound is supplied by Option Base 1.
{
    const ev = evalVBASingle(`
    Option Base 1
    Public mixedNpv, positiveError, negativeError, optionBaseError
    Sub Test()
        Dim mixed(1 To 3) As Double
        mixed(1) = -100 : mixed(2) = 40 : mixed(3) = 80
        mixedNpv = NPV(0.1, mixed)

        Dim positive(1 To 2) As Double
        positive(1) = 100 : positive(2) = 20
        On Error Resume Next
        Dim ignored As Double
        ignored = NPV(0.1, positive)
        positiveError = Err.Number
        Err.Clear

        Dim negative(1 To 2) As Double
        negative(1) = -100 : negative(2) = -20
        ignored = NPV(0.1, negative)
        negativeError = Err.Number
        Err.Clear

        Dim optionBaseMixed(1 To 3) As Double
        optionBaseMixed(1) = -100 : optionBaseMixed(2) = 40 : optionBaseMixed(3) = 80
        ignored = NPV(0.1, optionBaseMixed)
        optionBaseError = Err.Number
    End Sub
    `);
    ev.callProcedure('Test', []);
    assert.ok(Number.isFinite(ev.env.get('mixednpv')), 'NPV accepts mixed-sign cash flows');
    assert.strictEqual(ev.env.get('positiveerror'), 5,
        'NPV positive-only cash flows follow the VBA sign contract');
    assert.strictEqual(ev.env.get('negativeerror'), 5,
        'NPV negative-only cash flows follow the VBA sign contract');
    assert.strictEqual(ev.env.get('optionbaseerror'), 0,
        'NPV accepts mixed-sign Option Base 1 cash flows');
}
console.log('[PASS] NPVの正負キャッシュフロー契約');

// Bug 181-A: financial functions must reject invalid domains instead of leaking
// Infinity/NaN or silently returning a value.
{
    const ev = evalVBASingle(`
    Function ProbeNpv() As Long
        On Error Resume Next
        Dim ignored As Double: ignored = NPV(-1, Array(100, 200)): ProbeNpv = Err.Number
    End Function
    Function ProbePv() As Long
        On Error Resume Next
        Dim ignored As Double: ignored = PV(-1, 10, -100): ProbePv = Err.Number
    End Function
    Function ProbeSln() As Long
        On Error Resume Next
        Dim ignored As Double: ignored = SLN(1000, 100, 0): ProbeSln = Err.Number
    End Function
    Function ProbeSyd() As Long
        On Error Resume Next
        Dim ignored As Double: ignored = SYD(1000, 100, 0, 1): ProbeSyd = Err.Number
    End Function
    Function ProbeDdb() As Long
        On Error Resume Next
        Dim ignored As Double: ignored = DDB(1000, 100, 0, 1): ProbeDdb = Err.Number
    End Function
    Function ProbeNper() As Long
        On Error Resume Next
        Dim ignored As Double: ignored = NPer(0.1, 0, 1000): ProbeNper = Err.Number
    End Function
    Function ProbeRate() As Long
        On Error Resume Next
        Dim ignored As Double: ignored = Rate(0, -100, 100): ProbeRate = Err.Number
    End Function
    `);
    for (const name of ['ProbeNpv', 'ProbePv', 'ProbeSln', 'ProbeSyd', 'ProbeDdb', 'ProbeNper', 'ProbeRate']) {
        assert.strictEqual(ev.callProcedure(name, []), 5, `${name} is Error 5`);
    }
}
console.log('[PASS] Bug 181-A: 財務関数の不正引数をエラー化');

// SYD rejects a negative salvage value before the arithmetic result is classified.
{
    const ev = evalVBASingle(String.raw`
    Function ProbeSydNegativeSalvage() As Long
        On Error Resume Next
        Dim ignored As Double: ignored = SYD(100, -1, 2, 1): ProbeSydNegativeSalvage = Err.Number
    End Function
    `);
    assert.strictEqual(ev.callProcedure('ProbeSydNegativeSalvage', []), 5,
        'SYD negative salvage is Error 5');
}
console.log('[PASS] SYD negative salvage domain');

// Excel keeps this SLN result finite by scaling each operand before subtraction.
{
    const ev = evalVBASingle(`
    Public slnHuge As Variant, slnHugeErr As Long
    Sub Test()
        On Error Resume Next
        slnHuge = SLN(1E+308, -1E+308, 2)
        slnHugeErr = Err.Number
    End Sub
    `);
    ev.callProcedure('Test', []);
    assert.strictEqual(ev.env.get('slnhugeerr'), 0, 'SLN huge operands remain finite');
    assert.strictEqual(ev.env.get('slnhuge'), 1e308, 'SLN huge result matches Excel');
}
console.log('[PASS] SLN huge operand stability');

// SYD must not overflow an intermediate multiplication when the final
// mathematical result remains finite.
{
    const ev = evalVBASingle(String.raw`
    Public sydStable As Variant, sydStableErr As Long
    Sub Test()
        On Error Resume Next
        sydStable = SYD(1E+308, 0, 2, 1)
        sydStableErr = Err.Number
    End Sub
    `);
    ev.callProcedure('Test', []);
    assert.strictEqual(ev.env.get('sydstableerr'), 0,
        'SYD finite result does not become an intermediate overflow');
    assert.ok(Number.isFinite(ev.env.get('sydstable')),
        'SYD result remains finite');
}
console.log('[PASS] SYD intermediate overflow stability');

// Excel rejects a non-finite intermediate book/salvage difference in DDB.
{
    const ev = evalVBASingle(`
    Public ddbHugeErr As Long
    Sub Test()
        On Error Resume Next
        Dim ignored As Double
        ignored = DDB(1E+308, -1E+308, 2, 1)
        ddbHugeErr = Err.Number
    End Sub
    `);
    ev.callProcedure('Test', []);
    assert.strictEqual(ev.env.get('ddbhugeerr'), 5, 'DDB huge intermediate -> Error 5');
}
console.log('[PASS] DDB huge intermediate validation');

console.log('\n✅ Financial Functions: 全テスト通過');
