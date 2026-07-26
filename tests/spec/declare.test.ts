import { evalVBASingle, assert } from '../../test-libs/test-runner';

console.log('[Test Suite] Declare Statement (External Function) の検証');

const code = `
    Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long
    Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

    Function TestDeclare()
        Dim t
        t = GetTickCount()
        Sleep 100
        TestDeclare = t
    End Function
`;

const ev = evalVBASingle(code);

// Should not throw and should call the stub
assert.strictEqual(ev.callProcedure('TestDeclare', []), 0, 'Declare function should return 0 (stub default)');

// Declare stubs still enforce the declared arity, like ordinary VBA procedures.
const mismatch = evalVBASingle(`
    Declare Function OneArg Lib "kernel32" (ByVal value As Long) As Long
    Sub CallWithTooMany()
        OneArg 1, 2
    End Sub
`);
assert.throws(() => mismatch.callProcedure('CallWithTooMany', []), /Wrong number of arguments/,
    'Declare stub should reject too many arguments');

console.log('✅ Declare Statement: 全テスト通過');
