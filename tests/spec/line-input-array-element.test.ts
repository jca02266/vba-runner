import { evalVBASingle, assert } from '../../test-libs/test-runner';

const ev = evalVBASingle(`
    Function TestLineInputArrayElement() As String
        Dim lines(0 To 1) As String
        Open "lines.txt" For Output As #1
        Print #1, "first"
        Print #1, "second"
        Close #1
        Open "lines.txt" For Input As #1
        Line Input #1, lines(0)
        Line Input #1, lines(1)
        Close #1
        TestLineInputArrayElement = lines(0) & ":" & lines(1)
    End Function
`);

assert.strictEqual(ev.callProcedure('TestLineInputArrayElement', []), 'first:second');

console.log('✅ Line Input writes array elements');
