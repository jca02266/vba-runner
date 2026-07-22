import { evalVBASingle, assert } from '../../test-libs/test-runner';
import { MemoryFileSystem } from '../../src/engine/filesystem';

function readBytes(fs: MemoryFileSystem, file: string): number[] {
    const fd = fs.openSync(file, 'r');
    const buffer = new Uint8Array(fs.statSync(file).size);
    fs.readSync(fd, buffer, 0, buffer.length, 0);
    fs.closeSync(fd);
    return Array.from(buffer);
}

{
    const fs = new MemoryFileSystem();
    const ev = evalVBASingle(`
        Type BinaryRecord
            Id As Long
            Count As Integer
            Code As String * 3
        End Type

        Public longRead As Long
        Public textRead As String
        Public recordRead As BinaryRecord

        Sub RoundTrip()
            Dim n As Long
            Dim r As BinaryRecord
            Dim text As String

            n = 1234567
            r.Id = &H1020304
            r.Count = &H1122
            r.Code = "XYZ"
            text = "Aあ"

            Open "binary.dat" For Binary As #1
            Put #1, , n
            Put #1, , r
            Close #1

            Open "text.dat" For Binary As #1
            Put #1, , text
            Close #1

            Open "binary.dat" For Binary As #1
            Get #1, 1, longRead
            Get #1, 5, recordRead
            Close #1

            textRead = String$(Len(text), vbNullChar)
            Open "text.dat" For Binary As #1
            Get #1, 1, textRead
            Close #1
        End Sub
    `, { fs, sandboxRoot: '/sandbox' });

    ev.callProcedure('RoundTrip', []);

    assert.deepStrictEqual(readBytes(fs, '/sandbox/binary.dat'), [
        0x87, 0xd6, 0x12, 0x00,
        0x04, 0x03, 0x02, 0x01, 0x22, 0x11, 0x58, 0x59, 0x5a,
    ]);
    assert.deepStrictEqual(readBytes(fs, '/sandbox/text.dat'), [0x41, 0x82, 0xa0]);
    assert.strictEqual(ev.env.get('longRead'), 1234567);
    assert.strictEqual(ev.env.get('textRead'), 'Aあ');
    assert.strictEqual(ev.env.get('recordRead').id, 16909060);
    assert.strictEqual(ev.env.get('recordRead').count, 4386);
    assert.strictEqual(ev.env.get('recordRead').code, 'XYZ');
}

console.log('✅ Binary Put/Get: scalar, fixed-string UDT, and CP932 string round trips pass');

{
    const fs = new MemoryFileSystem();
    const ev = evalVBASingle(`
        Sub WriteConvertedByte()
            Open "byte.dat" For Binary As #1
            Put #1, , CByte(127)
            Close #1
        End Sub
    `, { fs, sandboxRoot: '/sandbox' });

    ev.callProcedure('WriteConvertedByte', []);
    assert.deepStrictEqual(readBytes(fs, '/sandbox/byte.dat'), [0x7f]);
}

console.log('✅ Binary Put accepts typed conversion expressions');
