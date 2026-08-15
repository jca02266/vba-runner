import { strict as assert } from 'node:assert';
import { evalVBAModules, evalVBASingle } from '../../test-libs/test-runner';

const evaluator = evalVBASingle(String.raw`
Option Compare Text

Public Function MatchAccentedRange() As String
    Select Case "À"
        Case "A" To "E"
            MatchAccentedRange = "range"
        Case Else
            MatchAccentedRange = "else"
    End Select
End Function
`);

assert.equal(evaluator.callProcedure('MatchAccentedRange', []), 'range');
console.log('[PASS] Select Case Option Compare Text non-ASCII range');

const matrix = evalVBAModules([
    {
        name: 'TextModule',
        code: String.raw`
Option Compare Text
Public Function Match(ByVal value As String) As String
    Select Case value
        Case "A" To "E": Match = "hit"
        Case Else: Match = "else"
    End Select
End Function
`,
    },
    {
        name: 'BinaryModule',
        code: String.raw`
Option Compare Binary
Public Function Match(ByVal value As String) As String
    Select Case value
        Case "A" To "E": Match = "hit"
        Case Else: Match = "else"
    End Select
End Function
`,
    },
]);

const values = ['A', 'a', 'À', 'à', 'Ê', 'ê', 'E', 'Z'];
const textResults = values.map((value) => matrix.callProcedure('TextModule.Match', [value]));
const binaryResults = values.map((value) => matrix.callProcedure('BinaryModule.Match', [value]));
assert.deepEqual(textResults, ['hit', 'hit', 'hit', 'hit', 'else', 'else', 'hit', 'else']);
assert.deepEqual(binaryResults, ['hit', 'else', 'else', 'else', 'else', 'else', 'hit', 'else']);
console.log('[PASS] Select Case Option Compare Text/Binary range matrix');
