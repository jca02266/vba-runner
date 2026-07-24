import { evalVBASingle, assert } from '../../test-libs/test-runner';
import { MemoryFileSystem } from '../../src/engine/filesystem';

const fs = new MemoryFileSystem();
const ev = evalVBASingle(`
    Function RoundTripBytes() As String
        Dim written(1 To 3) As Byte
        Dim readBack(1 To 3) As Byte
        written(1) = 1: written(2) = 2: written(3) = 255
        Open "bytes.bin" For Binary As #1
        Put #1, , written
        RoundTripBytes = CStr(LOF(1))
        Close #1
        Open "bytes.bin" For Binary As #1
        Get #1, , readBack
        Close #1
        RoundTripBytes = RoundTripBytes & ":" & CStr(readBack(1)) & ":" & CStr(readBack(2)) & ":" & CStr(readBack(3))
    End Function
`, { fs, sandboxRoot: '/sandbox' });

assert.strictEqual(ev.callProcedure('RoundTripBytes', []), '3:1:2:255');
console.log('✅ Binary Put/Get supports Byte arrays');
