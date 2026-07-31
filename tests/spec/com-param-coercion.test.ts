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

const fsoNamedResult = evalVBASingle(`
    Function ProbeFsoNamed() As String
        Dim fso As Object
        Set fso = CreateObject("Scripting.FileSystemObject")
        ProbeFsoNamed = CStr(fso.FileExists(filespec:="missing")) & "|" & _
            fso.GetBaseName(path:="C:\\x.txt")
    End Function
`);
assert.strictEqual(fsoNamedResult.callProcedure('ProbeFsoNamed', []), 'False|x',
    'FSO methods accept their documented named arguments');
assert.throws(
    () => evalVBASingle(`
        Dim fso As Object
        Set fso = CreateObject("Scripting.FileSystemObject")
        fso.FileExists Null
    `),
    'FSO required path arguments reject Null',
);

const streamNamedResult = evalVBASingle(`
    Function ProbeStreamNamed() As String
        Dim stream As Object
        Set stream = CreateObject("ADODB.Stream")
        stream.Open Options:=0
        stream.Write Buffer:="x"
        stream.Position = 0
        ProbeStreamNamed = stream.Read(NumBytes:=1)
    End Function
`);
assert.strictEqual(streamNamedResult.callProcedure('ProbeStreamNamed', []), 'x',
    'ADODB.Stream methods accept their documented named arguments');
assert.throws(
    () => evalVBASingle(`
        Dim stream As Object
        Set stream = CreateObject("ADODB.Stream")
        stream.Open
        stream.Write Null
    `),
    'ADODB.Stream.Write rejects Null Buffer',
);

const textStreamResult = evalVBASingle(`
    Function ProbeTextStreamNamed() As String
        Dim fso As Object, stream As Object
        Set fso = CreateObject("Scripting.FileSystemObject")
        Set stream = fso.CreateTextFile("textstream.txt", True, False)
        stream.WriteLine Text:="alpha"
        stream.Write Text:="beta"
        stream.Close
        Set stream = fso.OpenTextFile("textstream.txt", 1, False, -2)
        stream.Skip Characters:=2
        ProbeTextStreamNamed = stream.Read(Characters:=3)
        stream.Close
    End Function
`);
assert.strictEqual(textStreamResult.callProcedure('ProbeTextStreamNamed', []), 'pha',
    'TextStream methods accept documented named arguments');

const streamStateResult = evalVBASingle(`
    Function ProbeStreamState() As String
        Dim stream As Object
        Set stream = CreateObject("ADODB.Stream")
        stream.Open
        stream.Write Buffer:="abc"
        stream.Position = 0
        ProbeStreamState = stream.ReadText(NumChars:=2) & "|" & _
            CStr(stream.Position) & "|" & CStr(stream.Size)
    End Function
`);
assert.strictEqual(streamStateResult.callProcedure('ProbeStreamState', []), 'ab|2|3',
    'ADODB.Stream ReadText length and state are preserved');

const writePositionResult = evalVBASingle(`
    Function ProbeStreamWritePosition() As String
        Dim stream As Object
        Set stream = CreateObject("ADODB.Stream")
        stream.Open
        stream.WriteText "abc"
        ProbeStreamWritePosition = CStr(stream.Position) & "|" & CStr(stream.Size)
    End Function
`);
assert.strictEqual(writePositionResult.callProcedure('ProbeStreamWritePosition', []), '3|3',
    'ADODB.Stream writes advance Position');
assert.throws(
    () => evalVBASingle(`
        Dim stream As Object
        Set stream = CreateObject("ADODB.Stream")
        stream.Open
        stream.Position = Null
    `),
    'ADODB.Stream Position rejects Null',
);

assert.throws(
    () => evalVBASingle(`
        Dim stream As Object
        Set stream = CreateObject("ADODB.Stream")
        stream.Open
        stream.Close
        stream.ReadText
    `),
    'ADODB.Stream rejects ReadText after Close',
);

const textStreamProperties = evalVBASingle(`
    Function ProbeTextStreamProperties() As String
        Dim fso As Object, stream As Object
        Set fso = CreateObject("Scripting.FileSystemObject")
        Set stream = fso.CreateTextFile("textstream-properties.txt", True, False)
        stream.WriteLine "abc"
        stream.Close
        Set stream = fso.OpenTextFile("textstream-properties.txt", 1, False, -2)
        ProbeTextStreamProperties = CStr(stream.Line) & "|" & CStr(stream.Column) & "|" & _
            CStr(stream.AtEndOfLine) & "|" & CStr(stream.AtEndOfStream)
    End Function
`);
assert.strictEqual(textStreamProperties.callProcedure('ProbeTextStreamProperties', []), '1|1|False|False',
    'TextStream exposes line, column, and line-end state');

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
