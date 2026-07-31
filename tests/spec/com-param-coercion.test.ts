import { evalVBASingle, assert } from '../../test-libs/test-runner';

const result = evalVBASingle(`
    Function ProbeHostFlag() As String
        Dim host As Object
        Set host = CreateObject("Test.FlagHost")
        ProbeHostFlag = host.SetFlag("True") & "," & host.SetFlag("0")
    End Function
`, {
    setup: (ev) => {
        ev.registerComObject(() => ({
            __progId__: 'Test.FlagHost',
            __vbaParamSpecs__: {
                setflag: [{ name: 'Value', coerce: 'boolean' }],
            },
            setflag: (value: any) => value.value,
        } as any));
    },
});

assert.strictEqual(result.callProcedure('ProbeHostFlag', []), '-1,0',
    'generic COM metadata uses shared VBA Boolean coercion');
console.log('[PASS] generic COM parameter coercion');

const streamResult = evalVBASingle(`
    Function ProbeStreamNull() As Long
        Dim stream As Object
        Set stream = CreateObject("ADODB.Stream")
        stream.Open
        On Error Resume Next
        stream.WriteText Null
        ProbeStreamNull = Err.Number
    End Function
`);
assert.strictEqual(streamResult.callProcedure('ProbeStreamNull', []), 94,
    'COM string parameters reject Null with Invalid use of Null');

const dictionaryResult = evalVBASingle(`
    Function ProbeDictionaryMissing() As String
        Dim dictionary As Object
        Set dictionary = CreateObject("Scripting.Dictionary")
        dictionary.Add "1", "one"
        ProbeDictionaryMissing = CStr(dictionary.Item(1)) & "|" & CStr(dictionary.Count)
    End Function
`);
assert.strictEqual(dictionaryResult.callProcedure('ProbeDictionaryMissing', []), '|2',
    'Dictionary missing Item auto-creates VBA Empty');

const missingFileResult = evalVBASingle(`
    Function ProbeStreamMissingFile() As Long
        Dim stream As Object
        Set stream = CreateObject("ADODB.Stream")
        stream.Open
        On Error Resume Next
        stream.LoadFromFile "C:\\missing.bin"
        ProbeStreamMissingFile = Err.Number
    End Function
`);
assert.strictEqual(missingFileResult.callProcedure('ProbeStreamMissingFile', []), 53,
    'ADODB.Stream maps missing LoadFromFile to File not found');

const dictionaryNullResult = evalVBASingle(`
    Function ProbeDictionaryNull() As String
        Dim dictionary As Object
        Set dictionary = CreateObject("Scripting.Dictionary")
        On Error Resume Next
        dictionary.Add Null, "value"
        ProbeDictionaryNull = CStr(Err.Number) & "|" & CStr(dictionary.Count)
        Err.Clear
        dictionary.Remove Null
        ProbeDictionaryNull = ProbeDictionaryNull & "|" & CStr(Err.Number)
    End Function
`);
assert.strictEqual(dictionaryNullResult.callProcedure('ProbeDictionaryNull', []), '94|0|94',
    'Dictionary Null keys raise Invalid use of Null');

const dictionaryAssignmentResult = evalVBASingle(`
    Function ProbeDictionaryNullAssignment() As String
        Dim dictionary As Object
        Set dictionary = CreateObject("Scripting.Dictionary")
        On Error Resume Next
        dictionary.Item(Null) = "value"
        ProbeDictionaryNullAssignment = CStr(Err.Number) & "|" & CStr(dictionary.Count)
    End Function
`);
assert.strictEqual(dictionaryAssignmentResult.callProcedure('ProbeDictionaryNullAssignment', []), '94|0',
    'Dictionary Item assignment rejects Null keys');

const dictionaryCompareResult = evalVBASingle(`
    Function ProbeDictionaryCompareMode() As String
        Dim dictionary As Object
        Set dictionary = CreateObject("Scripting.Dictionary")
        dictionary.CompareMode = 1
        dictionary.Add "Key", "value"
        On Error Resume Next
        dictionary.Add "key", "duplicate"
        ProbeDictionaryCompareMode = CStr(Err.Number) & "|" & CStr(dictionary.Exists("KEY"))
    End Function
`);
assert.strictEqual(dictionaryCompareResult.callProcedure('ProbeDictionaryCompareMode', []), '457|True',
    'Dictionary Add normalizes keys through CompareMode');

const dictionaryReadResult = evalVBASingle(`
    Function ProbeDictionaryNullRead() As String
        Dim dictionary As Object
        Set dictionary = CreateObject("Scripting.Dictionary")
        On Error Resume Next
        Dim value As Variant
        value = dictionary(Null)
        ProbeDictionaryNullRead = CStr(Err.Number) & "|" & CStr(dictionary.Count)
    End Function
`);
assert.strictEqual(dictionaryReadResult.callProcedure('ProbeDictionaryNullRead', []), '94|0',
    'Dictionary default Item reads reject Null keys');

assert.throwsMatch(
    () => evalVBASingle(`
        Dim xhr As Object
        Set xhr = CreateObject("MSXML2.XMLHTTP")
        xhr.Open bstrMethod:=Null, bstrUrl:="http://example"
    `),
    /Invalid use of Null|error '94'/,
    'XMLHTTP.Open rejects Null method',
);
assert.throwsMatch(
    () => evalVBASingle(`
        Dim xhr As Object
        Set xhr = CreateObject("MSXML2.XMLHTTP")
        xhr.SetRequestHeader bstrHeader:="X-Test", bstrValue:=Null
    `),
    /Invalid use of Null|error '94'/,
    'XMLHTTP.SetRequestHeader rejects Null value',
);
evalVBASingle(`
    Dim xhr As Object
    Set xhr = CreateObject("MSXML2.XMLHTTP")
    xhr.Open bstrMethod:="GET", bstrUrl:="http://example", varAsync:=False
    xhr.Send varBody:=Null
`);
console.log('[PASS] XMLHTTP parameter metadata boundaries');
