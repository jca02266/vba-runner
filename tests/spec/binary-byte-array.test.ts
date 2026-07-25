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

const matrixFs = new MemoryFileSystem();
const matrixEv = evalVBASingle(`
    Function RoundTripMatrix() As String
        Dim written(0 To 1, 0 To 1) As Byte
        Dim readBack(0 To 1, 0 To 1) As Byte
        written(0, 0) = 1: written(0, 1) = 2
        written(1, 0) = 3: written(1, 1) = 4
        Open "matrix.bin" For Binary As #1
        Put #1, , written
        RoundTripMatrix = CStr(LOF(1))
        Close #1
        Open "matrix.bin" For Binary As #1
        Get #1, , readBack
        Close #1
        RoundTripMatrix = RoundTripMatrix & ":" & CStr(readBack(0, 0)) & ":" & CStr(readBack(0, 1)) & ":" & CStr(readBack(1, 0)) & ":" & CStr(readBack(1, 1))
    End Function
`, { fs: matrixFs, sandboxRoot: '/sandbox' });

assert.strictEqual(matrixEv.callProcedure('RoundTripMatrix', []), '4:1:2:3:4');
console.log('✅ Binary Put/Get supports multidimensional Byte arrays');

const intFs = new MemoryFileSystem();
const intEv = evalVBASingle(`
    Function RoundTripIntegers() As String
        Dim written(0 To 1) As Integer
        Dim readBack(0 To 1) As Integer
        written(0) = -1: written(1) = 300
        Open "integers.bin" For Binary As #1
        Put #1, , written
        RoundTripIntegers = CStr(LOF(1))
        Close #1
        Open "integers.bin" For Binary As #1
        Get #1, , readBack
        Close #1
        RoundTripIntegers = RoundTripIntegers & ":" & CStr(readBack(0)) & ":" & CStr(readBack(1))
    End Function
`, { fs: intFs, sandboxRoot: '/sandbox' });
assert.strictEqual(intEv.callProcedure('RoundTripIntegers', []), '4:-1:300');
console.log('✅ Binary Put/Get supports Integer arrays');

const intMatrixFs = new MemoryFileSystem();
const intMatrixEv = evalVBASingle(`
    Function RoundTripIntegerMatrix() As String
        Dim written(0 To 1, 0 To 1) As Integer
        Dim readBack(0 To 1, 0 To 1) As Integer
        written(0, 0) = -1: written(0, 1) = 2
        written(1, 0) = 300: written(1, 1) = -400
        Open "integer-matrix.bin" For Binary As #1
        Put #1, , written
        RoundTripIntegerMatrix = CStr(LOF(1))
        Close #1
        Open "integer-matrix.bin" For Binary As #1
        Get #1, , readBack
        Close #1
        RoundTripIntegerMatrix = RoundTripIntegerMatrix & ":" & CStr(readBack(0, 0)) & ":" & CStr(readBack(0, 1)) & ":" & CStr(readBack(1, 0)) & ":" & CStr(readBack(1, 1))
    End Function
`, { fs: intMatrixFs, sandboxRoot: '/sandbox' });
assert.strictEqual(intMatrixEv.callProcedure('RoundTripIntegerMatrix', []), '8:-1:2:300:-400');
console.log('✅ Binary Put/Get supports multidimensional Integer arrays');

const longFs = new MemoryFileSystem();
const longEv = evalVBASingle(`
    Function RoundTripLongs() As String
        Dim written(0 To 1, 0 To 1) As Long
        Dim readBack(0 To 1, 0 To 1) As Long
        written(0, 0) = -1: written(0, 1) = 2
        written(1, 0) = 300000: written(1, 1) = -400000
        Open "longs.bin" For Binary As #1
        Put #1, , written
        RoundTripLongs = CStr(LOF(1))
        Close #1
        Open "longs.bin" For Binary As #1
        Get #1, , readBack
        Close #1
        RoundTripLongs = RoundTripLongs & ":" & CStr(readBack(0, 0)) & ":" & CStr(readBack(0, 1)) & ":" & CStr(readBack(1, 0)) & ":" & CStr(readBack(1, 1))
    End Function
`, { fs: longFs, sandboxRoot: '/sandbox' });
assert.strictEqual(longEv.callProcedure('RoundTripLongs', []), '16:-1:2:300000:-400000');
console.log('✅ Binary Put/Get supports multidimensional Long arrays');

const booleanFs = new MemoryFileSystem();
const booleanEv = evalVBASingle(`
    Function RoundTripBooleans() As String
        Dim written(0 To 1) As Boolean
        Dim readBack(0 To 1) As Boolean
        written(0) = True: written(1) = False
        Open "booleans.bin" For Binary As #1
        Put #1, , written
        RoundTripBooleans = CStr(LOF(1))
        Close #1
        Open "booleans.bin" For Binary As #1
        Get #1, , readBack
        Close #1
        RoundTripBooleans = RoundTripBooleans & ":" & CStr(readBack(0)) & ":" & CStr(readBack(1))
    End Function
`, { fs: booleanFs, sandboxRoot: '/sandbox' });
assert.strictEqual(booleanEv.callProcedure('RoundTripBooleans', []), '4:True:False');
console.log('✅ Binary Put/Get supports Boolean arrays');

const doubleFs = new MemoryFileSystem();
const doubleEv = evalVBASingle(`
    Function RoundTripDoubles() As String
        Dim written(0 To 1) As Double
        Dim readBack(0 To 1) As Double
        written(0) = 1.25: written(1) = -2.5
        Open "doubles.bin" For Binary As #1
        Put #1, , written
        RoundTripDoubles = CStr(LOF(1))
        Close #1
        Open "doubles.bin" For Binary As #1
        Get #1, , readBack
        Close #1
        RoundTripDoubles = RoundTripDoubles & ":" & CStr(readBack(0)) & ":" & CStr(readBack(1))
    End Function
`, { fs: doubleFs, sandboxRoot: '/sandbox' });
assert.strictEqual(doubleEv.callProcedure('RoundTripDoubles', []), '16:1.25:-2.5');
console.log('✅ Binary Put/Get supports Double arrays');

const fixedStringFs = new MemoryFileSystem();
const fixedStringEv = evalVBASingle(`
    Function RoundTripFixedStrings() As String
        Dim written(0 To 1) As String * 2
        Dim readBack(0 To 1) As String * 2
        written(0) = "A": written(1) = "BC"
        Open "fixed-strings.bin" For Binary As #1
        Put #1, , written
        RoundTripFixedStrings = CStr(LOF(1))
        Close #1
        Open "fixed-strings.bin" For Binary As #1
        Get #1, , readBack
        Close #1
        RoundTripFixedStrings = RoundTripFixedStrings & ":" & readBack(0) & ":" & readBack(1)
    End Function
`, { fs: fixedStringFs, sandboxRoot: '/sandbox' });
assert.strictEqual(fixedStringEv.callProcedure('RoundTripFixedStrings', []), '4:A :BC');
console.log('✅ Binary Put/Get supports fixed-length String arrays');

const udtArrayFs = new MemoryFileSystem();
const udtArrayEv = evalVBASingle(`
    Type Point
        X As Long
        Y As Integer
    End Type

    Function RoundTripPoints() As String
        Dim written(0 To 1) As Point
        Dim readBack(0 To 1) As Point
        written(0).X = 1: written(0).Y = 2
        written(1).X = 3: written(1).Y = 4
        Open "points.bin" For Binary As #1
        Put #1, , written
        RoundTripPoints = CStr(LOF(1))
        Close #1
        Open "points.bin" For Binary As #1
        Get #1, , readBack
        Close #1
        RoundTripPoints = RoundTripPoints & ":" & CStr(readBack(0).X) & ":" & CStr(readBack(0).Y) & ":" & CStr(readBack(1).X) & ":" & CStr(readBack(1).Y)
    End Function
`, { fs: udtArrayFs, sandboxRoot: '/sandbox' });
assert.strictEqual(udtArrayEv.callProcedure('RoundTripPoints', []), '12:1:2:3:4');
console.log('✅ Binary Put/Get supports UDT arrays');

const shortReadFs = new MemoryFileSystem();
const shortReadEv = evalVBASingle(`
    Function ReadPastEnd() As String
        Dim written As Integer
        Dim readBack(0 To 1) As Integer
        written = 1
        Open "short.bin" For Binary As #1
        Put #1, , written
        Get #1, 1, readBack
    End Function
`, { fs: shortReadFs, sandboxRoot: '/sandbox' });
assert.throwsMatch(() => shortReadEv.callProcedure('ReadPastEnd', []), /Run-time error '62'/,
    'Binary Get beyond EOF → Error 62');
console.log('✅ Binary Get reports Input past end of file');
