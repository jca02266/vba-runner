import { evalVBASingle, assert } from '../../test-libs/test-runner';

const ev = evalVBASingle(`
    Function TestRandomRecordLength() As String
        Dim value As Long
        Open "records.dat" For Random As #1 Len = 8
        Put #1, 2, 123456
        TestRandomRecordLength = CStr(LOF(1))
        Close #1

        Open "records.dat" For Random As #1 Len = 8
        Get #1, 2, value
        Close #1
        TestRandomRecordLength = TestRandomRecordLength & ":" & CStr(value)
    End Function
`);

assert.strictEqual(ev.callProcedure('TestRandomRecordLength', []), '12:123456');

console.log('✅ Random files use Len for Put and Get record offsets');
