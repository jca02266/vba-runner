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

// Excel reports Error 5 when FinanceRate=-1 makes a later negative cash-flow
// discount denominator zero.  The initial cash flow alone remains valid.
{
    const ev = evalVBASingle(String.raw`
    Function ProbeMirrFinanceBoundary() As Long
        On Error Resume Next
        Dim flows(0 To 2) As Double, ignored As Double
        flows(0) = -100
        flows(1) = -50
        flows(2) = 200
        ignored = MIRR(flows, -1, 0.1)
        ProbeMirrFinanceBoundary = Err.Number
    End Function
    `);
    assert.strictEqual(ev.callProcedure('ProbeMirrFinanceBoundary', []), 5,
        'MIRR FinanceRate=-1 with a later negative cash flow is Error 5');
}
console.log('[PASS] MIRR FinanceRate zero-denominator boundary');

// The initial negative cash flow is enough for Excel to reject FinanceRate=-1.
{
    const ev = evalVBASingle(String.raw`
    Function ProbeMirrInitialOnlyBoundary() As Long
        On Error Resume Next
        Dim flows(0 To 2) As Double, ignored As Double
        flows(0) = -100
        flows(1) = 50
        flows(2) = 75
        ignored = MIRR(flows, -1, 0.1)
        ProbeMirrInitialOnlyBoundary = Err.Number
    End Function
    `);
    assert.strictEqual(ev.callProcedure('ProbeMirrInitialOnlyBoundary', []), 5,
        'MIRR FinanceRate=-1 with no later negative cash flow is Error 5');
}
console.log('[PASS] MIRR FinanceRate=-1 initial-only boundary');

// VBA accumulates MIRR's signed NPV values through a native extended-precision
// recurrence before storing each completed total back into a Double.
{
    const ev = evalVBASingle(String.raw`
    Sub MirrMatrixHost()
    End Sub
    `);
    const mirr = ev.env.get('mirr') as (
        values: number[], financeRate: number, reinvestRate: number,
    ) => number;
    const probe = (values: number[], financeRate: number, reinvestRate: number): number | string => {
        try {
            return mirr(values, financeRate, reinvestRate);
        } catch (error) {
            return `ERR:${(error as { number?: number }).number ?? 'unknown'}`;
        }
    };
    const expectValue = (
        label: string,
        values: number[],
        financeRate: number,
        reinvestRate: number,
        expected: number,
    ): void => {
        const actual = probe(values, financeRate, reinvestRate);
        assert.strictEqual(typeof actual, 'number', `${label} returns Double`);
        assertClose(actual as number, expected, { message: label });
    };

    const growthRates = [-0.5, -0.75, -0.875, -0.9375, -0.999, -0.999999,
        -0.99999999, -0.9999999999];
    const growthExpected = [0.118033988749895, 0.0606601717798212,
        0.0307764064044151, 0.0155048005794951, 0.000249968757810137,
        2.50000002477435e-7, 2.4992732328144e-9, 2.25147456234254e-10];
    growthRates.forEach((rate, index) => {
        expectValue(`XL-238 CASE=${index}`, [-100, 50, 100], 0.1, rate,
            growthExpected[index]);
    });

    for (const scale of [2 ** -500, 2 ** -40, 1, 2 ** 500]) {
        const values = [-100 * scale, 50 * scale, 100 * scale];
        expectValue(`XL-239 normal scale=${scale}`, values, 0.1, 0.12,
            0.24899959967968);
        expectValue(`XL-239 near scale=${scale}`, values, 0.1, -0.9999999999,
            2.25147456234254e-10);
    }

    const periodCases: Array<[number[], number]> = [
        [[-100, 150], 0.500000000473316],
        [[-100, 50, 100], 2.25147456234254e-10],
        [[-100, 50, 0, 100], 2.35309105534043e-10],
        [[-100, 50, 0, 0, 100], 2.88179924368137e-10],
        [[-100, 50, 0, 0, 0, 100], 2.63979282877358e-10],
        [[-100, 50, 0, 0, 0, 0, 0, 100], 1.04898534303288e-10],
    ];
    periodCases.forEach(([values, expected], index) => {
        expectValue(`XL-240 CASE=${index + 1}`, values, 0.1, -0.9999999999, expected);
    });

    const positionCases: Array<[number[], number]> = [
        [[-100, -50, 0, 50, 100], -0.0894198563072709],
        [[-100, 50, -50, 0, 100], -0.0828353339356729],
        [[-100, 50, 0, -50, 100], -0.0766370851899743],
        [[-100, 50, 100, -50, 0], -0.999990766370472],
        [[50, -100, -50, 0, 100], -0.0674624098474518],
        [[50, -100, 0, 100, -50], -0.997009659738562],
        [[100, 50, 0, -50, -100], -0.999999973784854],
        [[100, -50, 0, 50, -100], -0.997425167180969],
    ];
    positionCases.forEach(([values, expected], index) => {
        expectValue(`XL-241 CASE=${index + 1}`, values, 0.1, -0.9999999999, expected);
    });

    const interactionCases: Array<[number, number, number | string]> = [
        [0.1, 0.12, 0.0236050824969549],
        [-0.5, 0.12, -0.079483591748411],
        [0.1, -0.5, -0.0492618848096521],
        [-0.9999999999, 0.12, -0.999321757692808],
        [0.1, -0.9999999999, -0.117412915946094],
        [-0.9999999999, -0.9999999999, -0.999415196436182],
        [-1.1, -0.9999999999, 'ERR:5'],
        [-0.9999999999, -1.1, -0.999425110276349],
    ];
    interactionCases.forEach(([financeRate, reinvestRate, expected], index) => {
        const label = `XL-242 CASE=${index + 1}`;
        if (typeof expected === 'string') {
            assert.strictEqual(probe([-100, -50, 50, 100], financeRate, reinvestRate),
                expected, label);
        } else {
            expectValue(label, [-100, -50, 50, 100], financeRate, reinvestRate, expected);
        }
    });

    const financeShapes = [
        [-100, 50, 75],
        [-100, -50, 200],
        [-100, 100, -20, 75],
    ];
    const financeRates = [-2, -1.1, -1.0000000001, -1, -0.9999999999, -0.5, 0.1];
    const financeExpected: Array<Array<number | string>> = [
        [0.140175425099138, 1, 0.177673606017982],
        [0.140175425099138, 'ERR:5', -0.546393858936442],
        [0.140175425099138, 'ERR:5', -0.999999538956345],
        ['ERR:5', 'ERR:5', 'ERR:5'],
        [0.140175425099138, -0.999979999999178, -0.999999538956345],
        [0.140175425099138, 0, 0.0287926560528617],
        [0.140175425099138, 0.172603939955858, 0.189252600286169],
    ];
    financeRates.forEach((financeRate, rateIndex) => {
        financeShapes.forEach((values, shapeIndex) => {
            const expected = financeExpected[rateIndex][shapeIndex];
            const label = `XL-220 RATE=${financeRate} SHAPE=${shapeIndex + 1}`;
            if (typeof expected === 'string') {
                assert.strictEqual(probe(values, financeRate, 0.1), expected, label);
            } else {
                expectValue(label, values, financeRate, 0.1, expected);
            }
        });
    });
}
console.log('[PASS] MIRR VBA extended-precision and finance-rate matrices');

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

// BUG-00538: Excel reports overflow for the maximum cancellation boundary
// used by NPer, even though the direct logarithmic quotient is finite.
{
    const ev = evalVBASingle(String.raw`Function ProbeNPerBoundary() As Long
        On Error Resume Next
        Dim value As Double
        value = NPer(0.5, -1E+308, 1E+308, 0, 0)
        ProbeNPerBoundary = Err.Number
    End Function`);
    assert.strictEqual(ev.callProcedure('ProbeNPerBoundary', []), 6,
        'NPer maximum cancellation boundary is an overflow');
}
console.log('[PASS] BUG-00538: NPer maximum cancellation boundary');

// BUG-00550: Excel distinguishes NPer's rate-domain errors from its
// intermediate overflow boundary.  In particular, rate = 0 is the linear
// formula, rate = -1 is Error 5, and the finite-large shape is Error 6.
{
    const ev = evalVBASingle(String.raw`Function ProbeNPerExcelBoundaries() As String
        Dim value As Double, text As String
        On Error Resume Next
        Err.Clear: value = NPer(0.5, -1E+308, 1E+307, 0, 0)
        text = text & "finite-large=" & Err.Number & ":" & CStr(value) & ";"
        Err.Clear: value = NPer(0, -1E+308, 1E+308, 0, 0)
        text = text & "zero-rate=" & Err.Number & ":" & CStr(value) & ";"
        Err.Clear: value = NPer(-1, -1E+308, 1E+308, 0, 0)
        text = text & "minus-one=" & Err.Number & ";"
        ProbeNPerExcelBoundaries = text
    End Function`);
    assert.strictEqual(ev.callProcedure('ProbeNPerExcelBoundaries', []),
        'finite-large=6:0;zero-rate=0:1;minus-one=5;',
        'NPer preserves Excel Error 5/6 and zero-rate boundaries');
}
console.log('[PASS] BUG-00550: NPer rate and overflow boundaries');

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

// Financial linear combinations must preserve finite cancellation results
// instead of overflowing an individual product first.
{
    const ev = evalVBASingle(String.raw`
    Public fvStable As Variant, pvStable As Variant, pmtStable As Variant
    Public fvStableErr As Long, pvStableErr As Long, pmtStableErr As Long
    Sub Test()
        Dim huge As Double, value As Double
        huge = 1E+308
        On Error Resume Next
        Err.Clear: value = FV(0.5, 2, -huge * 0.48, huge * 0.8, 0): fvStable = value: fvStableErr = Err.Number
        Err.Clear: value = PV(0.5, 2, huge * 0.8, -huge * 1.7, 0): pvStable = value: pvStableErr = Err.Number
        Err.Clear: value = PMT(0.5, 2, huge * 0.8, -huge * 1.7, 0): pmtStable = value: pmtStableErr = Err.Number
    End Sub
    `);
    ev.callProcedure('Test', []);
    assert.strictEqual(ev.env.get('fvstableerr'), 0, 'FV cancellation remains finite');
    assert.strictEqual(ev.env.get('pvstableerr'), 0, 'PV cancellation remains finite');
    assert.strictEqual(ev.env.get('pmtstableerr'), 0, 'PMT cancellation remains finite');
    assert.ok(Number.isFinite(ev.env.get('fvstable')));
    assert.ok(Number.isFinite(ev.env.get('pvstable')));
    assert.ok(Number.isFinite(ev.env.get('pmtstable')));
}
console.log('[PASS] Financial linear-combination stability');

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
