import { evalVBASingle, assert } from '../../test-libs/test-runner';

const source = String.raw`
Option Explicit

Public Function Probe() As String
    Dim one As LongLong, zero As LongLong, pointer As LongPtr
    one = CLngLng("1")
    zero = CLngLng("0")
    pointer = CLngPtr("1")
    Probe = CStr(CBool(one)) & "|" & CStr(CBool(zero)) & "|" & CStr(CBool(pointer))
End Function
`;

const ev = evalVBASingle(source);
assert.strictEqual(
    ev.callProcedure('Probe', []),
    'True|False|True',
    'CBoolはLongLong/LongPtrの0と非ゼロを数値式として変換する',
);

console.log('✅ CBool LongLong/LongPtr coercion passes');
