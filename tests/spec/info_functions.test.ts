import { evalVBASingle, assert } from '../../test-libs/test-runner';
import { vbaMissing, vbaTrue, vbaFalse } from '../../src/engine/evaluator';

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

const ev = evalVBASingle(code);

// IsMissing
assert.strictEqual(ev.callProcedure('TestMissing', [vbaMissing]), vbaTrue, 'IsMissing(vbaMissing) should be vbaTrue');
assert.strictEqual(ev.callProcedure('TestMissing', [10]), vbaFalse, 'IsMissing(10) should be vbaFalse');

// VarType
assert.strictEqual(ev.callProcedure('TestVarType', []), 2, 'VarType(10) should be vbInteger (2)');

// Constants
assert.strictEqual(ev.callProcedure('TestConstants', []), vbaTrue, 'VBA type constants should be available');

console.log('✅ Information Functions: 全テスト通過');
