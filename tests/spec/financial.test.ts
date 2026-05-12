/**
 * Financial Functions (§6.1.2.6) のテスト
 */
import { Lexer } from '../../src/compiler/lexer';
import { Parser } from '../../src/compiler/parser';
import { Evaluator } from '../../src/compiler/evaluator';
import { assert } from '../ts/test-runner';

function evalVBA(code: string): any {
    const tokens = new Lexer(code).tokenize();
    const ast = new Parser(tokens).parse();
    const ev = new Evaluator(console.log);
    ev.evaluate(ast);
    return ev;
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

console.log('\n✅ Financial Functions: 全テスト通過');
