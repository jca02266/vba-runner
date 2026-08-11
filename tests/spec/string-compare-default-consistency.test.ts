/** Compare省略時のOption Compareと関数間整合。 */
import { evalVBAModules, assert } from '../../test-libs/test-runner';

const ev = evalVBAModules([
    {
        name: 'TextDefaults',
        code: String.raw`Option Compare Text
Public Function TextProbe() As String
    TextProbe = CStr(StrComp("A", "a")) & "|" & Replace("aBa", "b", "x")
End Function
`,
    },
    {
        name: 'BinaryDefaults',
        code: String.raw`Option Compare Binary
Public Function BinaryProbe() As String
    BinaryProbe = CStr(StrComp("A", "a")) & "|" & Replace("aBa", "b", "x")
End Function
`,
    },
]);

assert.strictEqual(ev.callProcedure('TextProbe', []), '0|axa',
    'Text defaults apply consistently to StrComp and Replace');
assert.strictEqual(ev.callProcedure('BinaryProbe', []), '-1|aBa',
    'Binary defaults apply consistently to StrComp and Replace');
console.log('[PASS] String Compare default consistency matrix');
