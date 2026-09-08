import { assert, evalVBASingle } from '../../test-libs/test-runner';

const source = String.raw`
Sub WidthProbe()
    Open "width.dat" For Output As #1
    On Error Resume Next
    Width #1, Null
    Debug.Print "WIDTH=" & Err.Number
    Close #1
End Sub
Sub RandomProbe()
    On Error Resume Next
    Open "random.dat" For Random Access Read Write As #1 Len = Null
    Debug.Print "RANDOM=" & Err.Number
End Sub
Sub SpcProbe()
    On Error Resume Next
    Debug.Print Spc(Null)
    Debug.Print "SPC=" & Err.Number
End Sub
Sub IndexProbe()
    Dim values(0 To 1) As Long
    On Error Resume Next
    values(Null) = 1
    Debug.Print "INDEX=" & Err.Number
End Sub
Sub SeekProbe()
    Open "seek.dat" For Binary As #1
    On Error Resume Next
    Seek #1, Null
    Debug.Print "SEEK=" & Err.Number
End Sub
Sub TabProbe()
    On Error Resume Next
    Debug.Print Tab(Null)
    Debug.Print "TAB=" & Err.Number
End Sub
`;

const output: string[] = [];
const ev = evalVBASingle(source, { onPrint: (line: string) => output.push(line) });
for (const name of ['WidthProbe', 'RandomProbe', 'SpcProbe', 'IndexProbe', 'SeekProbe', 'TabProbe']) {
    ev.callProcedure(name, []);
}

assert.deepStrictEqual(output, [
    'WIDTH=94',
    'RANDOM=94',
    'SPC=94',
    'INDEX=94',
    'SEEK=94',
    'TAB=94',
]);
console.log('[PASS] Null numeric argument boundaries map to Invalid use of Null');
