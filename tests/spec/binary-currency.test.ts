import { evalVBASingle, assert } from '../../test-libs/test-runner';

const ev = evalVBASingle(`
    Function TestBinaryCurrency() As String
        Dim value As Currency
        value = CCur(-12.3456)
        Open "currency.bin" For Binary As #1
        Put #1, , value
        TestBinaryCurrency = CStr(LOF(1))
        Close #1
        value = 0
        Open "currency.bin" For Binary As #1
        Get #1, , value
        Close #1
        TestBinaryCurrency = TestBinaryCurrency & ":" & CStr(value)
    End Function
`);

assert.strictEqual(ev.callProcedure('TestBinaryCurrency', []), '8:-12.3456');

console.log('✅ Binary Put and Get support Currency');
