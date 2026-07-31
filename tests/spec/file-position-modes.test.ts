import { assert, evalVBASingle } from '../../test-libs/test-runner';

const random = evalVBASingle(`
    Function ProbeRandomPosition() As Long
        Dim value As Long, result As Long
        Open "position-random.dat" For Random As #1 Len = 4
        result = Loc(1) * 100 + Seek(1) * 10
        Seek #1, 2
        result = result + Loc(1) * 100 + Seek(1) * 10
        Put #1, 1, value
        result = result + Loc(1) * 100 + Seek(1) * 10
        Put #1, 2, value
        result = result + Loc(1) * 100 + Seek(1) * 10
        Get #1, 1, value
        result = result + Loc(1) * 100 + Seek(1) * 10
        Get #1, 2, value
        result = result + Loc(1) * 100 + Seek(1) * 10
        On Error Resume Next
        Get #1, 3, value
        If EOF(1) Then result = result + 1000
        result = result + Loc(1) * 100 + Seek(1) * 10 + Err.Number
        Close #1
        ProbeRandomPosition = result
    End Function
`);
assert.strictEqual(random.callProcedure('ProbeRandomPosition', []), 2022,
    'Random Loc/Seek use record numbers and EOF follows failed Get');

const binary = evalVBASingle(`
    Function ProbeBinaryPosition() As Long
        Dim value As Long, result As Long
        Open "position-binary.dat" For Binary As #1
        If EOF(1) Then result = result + 1
        Put #1, , value
        result = result + Loc(1) * 10 + Seek(1)
        If EOF(1) Then result = result + 100
        On Error Resume Next
        Get #1, , value
        Get #1, , value
        If EOF(1) Then result = result + 1000
        result = result + Err.Number
        Close #1
        ProbeBinaryPosition = result
    End Function
`);
assert.strictEqual(binary.callProcedure('ProbeBinaryPosition', []), 1107,
    'Binary EOF remains false after Put and becomes true after failed Get');

const sequential = evalVBASingle(`
    Function ProbeSequentialPosition() As Long
        Dim result As Long, value As String
        Open "position-sequential.dat" For Output As #1
        Print #1, "AB"
        result = Loc(1) * 100 + Seek(1)
        If EOF(1) Then result = result + 1000
        Close #1
        ProbeSequentialPosition = result
    End Function
`);
assert.strictEqual(sequential.callProcedure('ProbeSequentialPosition', []), 1005,
    'Sequential Loc uses byte position divided by 128');

console.log('[PASS] mode-specific EOF/Loc/Seek semantics');
