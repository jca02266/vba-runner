import { assert, evalVBASingle } from '../../test-libs/test-runner';

const ev = evalVBASingle(`
    Function CheckAppendPosition() As Long
        Dim result As Long
        Open "append-position.dat" For Output As #1
        Print #1, "AB"
        Close #1

        Open "append-position.dat" For Append As #1
        If EOF(1) Then result = result + 1000
        result = result + Seek(1)
        Print #1, "CD"
        If EOF(1) Then result = result + 100000
        result = result + Seek(1)
        Close #1
        CheckAppendPosition = result
    End Function
`);

// "AB" + CRLF is four bytes in the text encoding used by the file I/O layer.
// Append starts at that byte position and advances to eight after the second line.
assert.strictEqual(ev.callProcedure('CheckAppendPosition', []), 101014,
    'Append initializes and advances the evaluator file position at EOF');
console.log('[PASS] Append EOF/Seek position stays at the physical end');
