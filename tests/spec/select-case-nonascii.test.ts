import { strict as assert } from 'node:assert';
import { evalVBASingle } from '../../test-libs/test-runner';

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
