/**
 * Financial Functions (§6.1.2.6) のテスト
 */
import { evalVBASingle, assert } from '../../test-libs/test-runner';

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

// --- Bug 24-1: NPV が 1-based 配列で NaN を返す ---
// Dim flows(1 To N) で宣言した配列を渡すと vbaBase=1 のため
// values.map(Number) がインデックス 0 (undefined → NaN) を含んでいた
{
    const ev = evalVBASingle(`
    Public resNpv1, resNpv0
    Sub Test()
        Dim flows1(1 To 3) As Double
        flows1(1) = 30000 : flows1(2) = 40000 : flows1(3) = 50000
        resNpv1 = NPV(0.1, flows1)

        Dim flows0(2) As Double
        flows0(0) = 30000 : flows0(1) = 40000 : flows0(2) = 50000
        resNpv0 = NPV(0.1, flows0)
    End Sub
    `);
    ev.callProcedure('Test', []);
    const npv1 = ev.env.get('resnpv1') as number;
    const npv0 = ev.env.get('resnpv0') as number;
    assert.ok(!isNaN(npv1), `Bug 24-1: NPV 1-based should not be NaN (got ${npv1})`);
    assert.ok(Math.abs(npv1 - npv0) < 0.01, `NPV 1-based should equal 0-based (${npv1} vs ${npv0})`);
    assert.ok(Math.abs(npv1 - 97896.32) < 1, `NPV(0.1,[30000,40000,50000]) ≈ 97896.32 (got ${npv1.toFixed(2)})`);
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

console.log('\n✅ Financial Functions: 全テスト通過');
