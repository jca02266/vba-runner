import { evalVBASingle, assert } from '../../test-libs/test-runner';

const ev = evalVBASingle(`
    Function TestBinaryDate() As String
        Dim value As Date
        value = #2024/01/02 12:00:00#
        Open "date.bin" For Binary As #1
        Put #1, , value
        TestBinaryDate = CStr(LOF(1))
        Close #1
        value = #1900/01/01#
        Open "date.bin" For Binary As #1
        Get #1, , value
        Close #1
        TestBinaryDate = TestBinaryDate & ":" & Format(value, "yyyy/mm/dd hh:nn:ss")
    End Function
`);

assert.strictEqual(ev.callProcedure('TestBinaryDate', []), '8:2024/01/02 12:00:00');

console.log('✅ Binary Put and Get support Date');
