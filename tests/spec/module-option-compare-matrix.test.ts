/** モジュールごとのOption Compare継承境界。 */
import { evalVBAModules, assert } from '../../test-libs/test-runner';

const ev = evalVBAModules([
    {
        name: 'TextModule',
        code: String.raw`Option Compare Text
Public Function TextProbe(ByVal value As String) As Long
    TextProbe = StrComp("Alpha", value)
End Function
`,
    },
    {
        name: 'BinaryModule',
        code: String.raw`Option Compare Binary
Public Function BinaryProbe(ByVal value As String) As Long
    BinaryProbe = StrComp("Alpha", value)
End Function
`,
    },
]);

assert.strictEqual(ev.callProcedure('TextProbe', ['alpha']), 0,
    'Text module inherits Text comparison');
assert.strictEqual(ev.callProcedure('BinaryProbe', ['alpha']), -1,
    'Binary module inherits Binary comparison');
assert.strictEqual(ev.callProcedure('TextProbe', ['ALPHA']), 0,
    'Text module keeps case-insensitive matching');
assert.strictEqual(ev.callProcedure('BinaryProbe', ['ALPHA']), 1,
    'Binary module keeps case-sensitive matching');
console.log('[PASS] Module Option Compare matrix');
