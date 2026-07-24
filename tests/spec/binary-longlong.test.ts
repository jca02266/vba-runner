import { evalVBASingle, assert } from '../../test-libs/test-runner';
import { MemoryFileSystem } from '../../src/engine/filesystem';

const fs = new MemoryFileSystem();
const ev = evalVBASingle(`
    Function RoundTripLongLong() As String
        Dim written As LongLong
        Dim readBack As LongLong
        written = -1234567890123^

        Open "longlong.bin" For Binary As #1
        Put #1, , written
        RoundTripLongLong = CStr(LOF(1))
        Close #1

        Open "longlong.bin" For Binary As #1
        Get #1, , readBack
        Close #1

        RoundTripLongLong = RoundTripLongLong & ":" & CStr(readBack)
    End Function
`, { fs, sandboxRoot: '/sandbox' });

assert.strictEqual(ev.callProcedure('RoundTripLongLong', []), '8:-1234567890123');
console.log('✅ Binary Put/Get supports LongLong');

const ptrFs = new MemoryFileSystem();
const ptrEv = evalVBASingle(`
    Function RoundTripLongPtr() As String
        Dim written As LongPtr
        Dim readBack As LongPtr
        written = CLngPtr(123456789)

        Open "longptr.bin" For Binary As #1
        Put #1, , written
        RoundTripLongPtr = CStr(LOF(1))
        Close #1

        Open "longptr.bin" For Binary As #1
        Get #1, , readBack
        Close #1

        RoundTripLongPtr = RoundTripLongPtr & ":" & CStr(readBack)
    End Function
`, { fs: ptrFs, sandboxRoot: '/sandbox' });

assert.strictEqual(ptrEv.callProcedure('RoundTripLongPtr', []), '8:123456789');
console.log('✅ Binary Put/Get supports LongPtr');
