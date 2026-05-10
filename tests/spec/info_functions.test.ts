import { assert } from '../ts/test-runner';
import { Evaluator, vbaMissing, vbaTrue, vbaFalse } from '../../src/compiler/evaluator';
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

console.log('[Test Suite] Information Functions (IsMissing, VarType, etc.) の検証');

const code = `
    Function TestMissing(Optional x)
        TestMissing = IsMissing(x)
    End Function

    Function TestVarType()
        TestVarType = VarType(10)
    End Function

    Function TestConstants()
        If vbinteger = 2 And vbstring = 8 Then
            TestConstants = True
        Else
            TestConstants = False
        End If
    End Function
`;

const ev = evalVBA(code);

// IsMissing
assert.strictEqual(ev.callProcedure('TestMissing', [vbaMissing]), vbaTrue, 'IsMissing(vbaMissing) should be vbaTrue');
assert.strictEqual(ev.callProcedure('TestMissing', [10]), vbaFalse, 'IsMissing(10) should be vbaFalse');

// VarType
assert.strictEqual(ev.callProcedure('TestVarType', []), 5, 'VarType(10) should be vbDouble (5) in JS implementation');

// Constants
assert.strictEqual(ev.callProcedure('TestConstants', []), vbaTrue, 'VBA type constants should be available');

console.log('✅ Information Functions: 全テスト通過');
