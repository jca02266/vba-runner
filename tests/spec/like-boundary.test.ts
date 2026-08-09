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

const selectText = evalVBAModules([
    {
        name: 'SelectText',
        code: `Option Compare Text
Function SelectMatch() As String
    Select Case "abc"
        Case "ABC": SelectMatch = "T"
        Case Else: SelectMatch = "F"
    End Select
End Function`,
    },
]);
assert.equal(selectText.callProcedure('SelectMatch', []), 'T',
    'Select Case string comparison follows Option Compare Text');

const localeText = evalVBASingle(String.raw`Option Compare Text
Function AccentRange(value As String) As String
    AccentRange = IIf(value Like "[A-E]", "T", "F")
End Function`);
assert.equal(localeText.callProcedure('AccentRange', ['À']), 'T',
    'Option Compare Text Like ranges include À within [A-E]');
assert.equal(localeText.callProcedure('AccentRange', ['à']), 'T',
    'Option Compare Text Like ranges include à within [A-E]');
assert.equal(localeText.callProcedure('AccentRange', ['Ê']), 'F',
    'Option Compare Text Like ranges do not over-include Ê after E');

const invalidPatterns = evalVBASingle(String.raw`Option Compare Text
Function PatternError(ByVal pattern As String) As Long
    On Error Resume Next
    Err.Clear
    Dim matched As Boolean
    matched = "Y" Like pattern
    PatternError = Err.Number
End Function
Function EmptyPatternMatch(ByVal value As String) As String
    EmptyPatternMatch = IIf(value Like "[]", "T", "F")
End Function`);
assert.equal(invalidPatterns.callProcedure('PatternError', ['[Z-A]']), 93,
    'Like rejects descending ranges with Error 93');
assert.equal(invalidPatterns.callProcedure('PatternError', ['[ABC']), 93,
    'Like rejects unterminated character lists with Error 93');
assert.equal(invalidPatterns.callProcedure('EmptyPatternMatch', ['']), 'T',
    'Like [] matches the zero-length string');
assert.equal(invalidPatterns.callProcedure('EmptyPatternMatch', ['Y']), 'F',
    'Like [] does not match a non-empty string');

const classInvalidPatterns = evalVBAModules([
    {
        name: 'LikeProbe',
        parseAsClass: 'LikeProbe',
        code: String.raw`Public Function PatternError(ByVal pattern As String) As Long
    On Error Resume Next
    Dim matched As Boolean
    matched = "Y" Like pattern
    PatternError = Err.Number
End Function`,
    },
    {
        name: 'LikeClassCaller',
        code: String.raw`Function ClassPatternError(ByVal pattern As String) As Long
    Dim probe As New LikeProbe
    ClassPatternError = probe.PatternError(pattern)
End Function`,
    },
]);
assert.equal(classInvalidPatterns.callProcedure('ClassPatternError', ['[Z-A]']), 93,
    'Like Error 93 propagates through a class method');
