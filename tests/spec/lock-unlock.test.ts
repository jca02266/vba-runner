import { assert, evalVBASingle } from '../../test-libs/test-runner';

const evalVba = (source: string) => evalVBASingle(source);

{
    const ev = evalVba(`
        Sub LockThenUnlock()
            Dim f As Integer
            f = FreeFile
            Open "C:\\ledger\\stock.dat" For Binary As #f
            Lock #f, 3 To 6
            Unlock #f, 3 To 6
            Close #f
        End Sub
    `);
    ev.callProcedure('LockThenUnlock', []);
    console.log('[PASS] Lock/Unlock range');
}

{
    const ev = evalVba(`
        Sub LockClosedFile()
            Dim f As Integer
            f = FreeFile
            Open "C:\\ledger\\closed.dat" For Binary As #f
            Close #f
            Lock #f
        End Sub
    `);
    let error: any;
    try { ev.callProcedure('LockClosedFile', []); } catch (e) { error = e; }
    assert.strictEqual(error?.number, 52, 'Lock on a closed file raises Bad file name or number');
    console.log('[PASS] Lock closed file');
}

{
    const ev = evalVba(`
        Sub DoubleLock()
            Dim f As Integer
            f = FreeFile
            Open "C:\\ledger\\double.dat" For Binary As #f
            Lock #f, 3 To 6
            Lock #f, 3 To 6
        End Sub
    `);
    let error: any;
    try { ev.callProcedure('DoubleLock', []); } catch (e) { error = e; }
    assert.strictEqual(error?.number, 75, 'overlapping Lock raises Path/File access error');
    console.log('[PASS] Lock overlap');
}
