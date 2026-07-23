import { evalVBASingle, assert } from '../../test-libs/test-runner';
import { MemoryFileSystem } from '../../src/engine/filesystem';

function readBytes(fs: MemoryFileSystem, file: string): number[] {
    const fd = fs.openSync(file, 'r');
    const buffer = new Uint8Array(fs.statSync(file).size);
    fs.readSync(fd, buffer, 0, buffer.length, 0);
    fs.closeSync(fd);
    return Array.from(buffer);
}

const fs = new MemoryFileSystem();
const ev = evalVBASingle(`
    Type BinaryRecord
        Values(1 To 2) As Long
        Codes(0 To 1) As String * 2
    End Type

    Function RoundTrip() As String
        Dim written As BinaryRecord
        Dim readBack As BinaryRecord

        written.Values(1) = 10
        written.Values(2) = 20
        written.Codes(0) = "A"
        written.Codes(1) = "XY"

        Open "record.bin" For Binary As #1
        Put #1, , written
        RoundTrip = CStr(LOF(1))
        Close #1

        Open "record.bin" For Binary As #1
        Get #1, , readBack
        Close #1

        RoundTrip = RoundTrip & ":" & CStr(readBack.Values(1)) & ":" & _
            CStr(readBack.Values(2)) & ":" & readBack.Codes(0) & ":" & readBack.Codes(1)
    End Function
`, { fs, sandboxRoot: '/sandbox' });

assert.strictEqual(ev.callProcedure('RoundTrip', []), '12:10:20:A :XY');
assert.deepStrictEqual(readBytes(fs, '/sandbox/record.bin'), [
    10, 0, 0, 0,
    20, 0, 0, 0,
    0x41, 0x20,
    0x58, 0x59,
]);

console.log('✅ Binary Put/Get supports fixed-size UDT member arrays');
