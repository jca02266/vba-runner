import { assert } from '../ts/test-runner';
import { Evaluator, vbaTrue, vbaFalse, vbaNull } from '../../src/compiler/evaluator';
import { Parser } from '../../src/compiler/parser';
import { Lexer } from '../../src/compiler/lexer';

function evalVBA(code: string) {
    const tokens = new Lexer(code).tokenize();
    const parser = new Parser(tokens);
    const program = parser.parse();
    const evaluator = new Evaluator((s) => console.log(s));
    evaluator.evaluate(program);
    return evaluator;
}

console.log('[Test Suite] Choose and Switch Functions の検証');

const code = `
    Function TestChoose(idx)
        TestChoose = Choose(idx, "A", "B", "C")
    End Function

    Function TestSwitch(val)
        TestSwitch = Switch(val = 1, "One", val = 2, "Two", val > 2, "Many")
    End Function
`;

const ev = evalVBA(code);

// Choose
assert.strictEqual(ev.callProcedure('TestChoose', [1]), "A", 'Choose(1) should be "A"');
assert.strictEqual(ev.callProcedure('TestChoose', [2]), "B", 'Choose(2) should be "B"');
assert.strictEqual(ev.callProcedure('TestChoose', [4]), vbaNull, 'Choose(4) should be Null');

// Switch
assert.strictEqual(ev.callProcedure('TestSwitch', [1]), "One", 'Switch(1) should be "One"');
assert.strictEqual(ev.callProcedure('TestSwitch', [2]), "Two", 'Switch(2) should be "Two"');
assert.strictEqual(ev.callProcedure('TestSwitch', [5]), "Many", 'Switch(5) should be "Many"');

console.log('✅ Choose and Switch: 全テスト通過');
