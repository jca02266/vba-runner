import { evalVBASingle, assert } from '../../test-libs/test-runner';
import { MemoryFileSystem } from '../../src/engine/filesystem';

const fs = new MemoryFileSystem();
const ev = evalVBASingle(`
    Function RoundTripBoolean() As String
        Dim writtenTrue As Boolean
        Dim writtenFalse As Boolean
        Dim readTrue As Boolean
        Dim readFalse As Boolean
        writtenTrue = True
        writtenFalse = False

        Open "boolean.bin" For Binary As #1
        Put #1, , writtenTrue
        Put #1, , writtenFalse
        RoundTripBoolean = CStr(LOF(1))
        Close #1

        Open "boolean.bin" For Binary As #1
        Get #1, , readTrue
        Get #1, , readFalse
        Close #1

        RoundTripBoolean = RoundTripBoolean & ":" & CStr(readTrue) & ":" & CStr(readFalse)
    End Function
`, { fs, sandboxRoot: '/sandbox' });

assert.strictEqual(ev.callProcedure('RoundTripBoolean', []), '4:True:False');
console.log('✅ Binary Put/Get supports Boolean');
