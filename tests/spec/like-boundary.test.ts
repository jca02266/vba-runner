import { strict as assert } from 'node:assert';
import { evalVBAModules, evalVBASingle } from '../../test-libs/test-runner';

const textOnly = evalVBAModules([
    {
        name: 'TextModule',
        code: `Option Compare Text
Function TextMatch() As String
    TextMatch = IIf("ABC" Like "abc", "T", "F")
End Function`,
    },
    {
        name: 'BinaryModule',
        code: `Option Compare Binary
Function BinaryMatch() As String
    BinaryMatch = IIf("ABC" Like "abc", "T", "F")
End Function`,
    },
]);
assert.equal(textOnly.callProcedure('TextMatch', []), 'T', 'Option Compare Text is module-scoped');
assert.equal(textOnly.callProcedure('BinaryMatch', []), 'F', 'Option Compare Binary is module-scoped');

const classBoundary = evalVBASingle(`
Function LikeClosingBracket() As String
    LikeClosingBracket = IIf("]" Like "[]]", "T", "F")
End Function`);
assert.equal(classBoundary.callProcedure('LikeClosingBracket', []), 'T',
    'Like character class accepts a closing bracket as its first character');

console.log('✅ Like module comparison and character-class boundaries');
