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

console.log('\n✅ Financial Functions: 全テスト通過');
