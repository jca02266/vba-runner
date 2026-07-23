import { evalVBASingle, assert } from '../../test-libs/test-runner';

const ev = evalVBASingle(`
    Function TestWithDictionaryBang() As String
        Dim dict As Object
        Set dict = CreateObject("Scripting.Dictionary")
        With dict
            !Key = "value"
            TestWithDictionaryBang = !Key
        End With
    End Function
`);

assert.strictEqual(ev.callProcedure('TestWithDictionaryBang', []), 'value');

console.log('✅ With-block bang dictionary access works');
