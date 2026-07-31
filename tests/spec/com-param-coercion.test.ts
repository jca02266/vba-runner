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
