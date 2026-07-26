import { evalVBAModules, assert } from '../../test-libs/test-runner';

const ev = evalVBAModules([
    {
        name: 'KeyObject',
        parseAsClass: 'KeyObject',
        code: 'Public Id As Long',
    },
    {
        name: 'Main',
        code: `
            Function ObjectKeyRoundTrip() As String
                Dim d As Object
                Dim key As New KeyObject
                Set d = CreateObject("Scripting.Dictionary")
                key.Id = 7
                d.Add key, "value"
                ObjectKeyRoundTrip = CStr(d.Exists(key)) & ":" & d.Item(key)
            End Function
        `,
    },
]);

assert.strictEqual(ev.callProcedure('ObjectKeyRoundTrip', []), 'True:value');
console.log('✅ Dictionary accepts object keys and preserves object identity');
