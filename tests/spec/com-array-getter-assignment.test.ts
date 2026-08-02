import { evalVBASingle, assert } from '../../test-libs/test-runner';

const ev = evalVBASingle(`
Function ReadComArray() As Long
    Dim source As Object, target() As Long
    Set source = CreateObject("Test.Array")
    target = source.Values
    ReadComArray = target(0)
End Function
`);

ev.registerComObject(() => ({
    __progId__: 'Test.Array',
    Values: [7],
}) as any);

assert.strictEqual(ev.callProcedure('ReadComArray', []), 7,
    'COM mock SAFEARRAY getter can initialize a typed array');
console.log('[PASS] COM mock array getter boundary');
