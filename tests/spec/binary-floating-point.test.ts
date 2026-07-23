import { evalVBASingle, assert } from '../../test-libs/test-runner';

const ev = evalVBASingle(`
    Function TestBinaryFloatingPoint() As String
        Dim singleValue As Single
        Dim doubleValue As Double
        singleValue = 1.5
        doubleValue = 3.25
        Open "numbers.bin" For Binary As #1
        Put #1, , singleValue
        Put #1, , doubleValue
        TestBinaryFloatingPoint = CStr(LOF(1))
        Close #1

        singleValue = 0
        doubleValue = 0
        Open "numbers.bin" For Binary As #1
        Get #1, , singleValue
        Get #1, , doubleValue
        Close #1
        TestBinaryFloatingPoint = TestBinaryFloatingPoint & ":" & CStr(singleValue) & ":" & CStr(doubleValue)
    End Function
`);

assert.strictEqual(ev.callProcedure('TestBinaryFloatingPoint', []), '12:1.5:3.25');

console.log('✅ Binary Put and Get support Single and Double');
