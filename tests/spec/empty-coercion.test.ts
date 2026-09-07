import { evalVBASingle, assert } from '../../test-libs/test-runner';

const ev = evalVBASingle(String.raw`
Function Probe() As String
    Dim value As Variant
    value = Empty
    Probe = CStr(value)
End Function
`);

assert.strictEqual(ev.callProcedure('Probe', []), '', 'EmptyをCStrへ変換すると空文字列');
console.log('[PASS] Empty string coercion through procedure');
