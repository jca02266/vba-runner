import { evalVBASingle, assert } from '../../test-libs/test-runner';
import { vbaNull } from '../../src/engine/evaluator';

console.log('[Test Suite] Choose and Switch Functions の検証');

const code = `
    Function TestChoose(idx)
        TestChoose = Choose(idx, "A", "B", "C")
    End Function

    Function TestSwitch(val)
        TestSwitch = Switch(val = 1, "One", val = 2, "Two", val > 2, "Many")
    End Function
`;

const ev = evalVBASingle(code);

// Choose
assert.strictEqual(ev.callProcedure('TestChoose', [1]), "A", 'Choose(1) should be "A"');
assert.strictEqual(ev.callProcedure('TestChoose', [2]), "B", 'Choose(2) should be "B"');
assert.strictEqual(ev.callProcedure('TestChoose', [4]), vbaNull, 'Choose(4) should be Null');

// Switch
assert.strictEqual(ev.callProcedure('TestSwitch', [1]), "One", 'Switch(1) should be "One"');
assert.strictEqual(ev.callProcedure('TestSwitch', [2]), "Two", 'Switch(2) should be "Two"');
assert.strictEqual(ev.callProcedure('TestSwitch', [5]), "Many", 'Switch(5) should be "Many"');

// Switch: 奇数個引数（不対ペア）は実行時エラー
// spec §6.1.2.8.1.16: "If the parts aren't properly paired, a run-time error occurs."
{
    const ev = evalVBASingle('');
    assert.throwsMatch(
        () => ev.evalExpression('Switch(False, "A", True)'),
        /Invalid procedure call|argument/i,
        'Switch with odd args should error',
    );
    console.log('[PASS] Switch 奇数引数はエラー');
}

console.log('✅ Choose and Switch: 全テスト通過');
