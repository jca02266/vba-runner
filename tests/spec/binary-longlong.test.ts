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

// Bug 208-A: LongLong arithmetic must not pass through IEEE-754 Number.
const arithmeticEv = evalVBASingle(`
    Function LongLongArithmetic() As String
        Dim x As LongLong
        x = CLngLng("9007199254740993")
        LongLongArithmetic = CStr(x + 1) & "|" & CStr(x - 1) & "|" & _
            CStr(x * 3) & "|" & CStr(x \\ 3) & "|" & CStr(x Mod 3) & "|" & TypeName(x + 1)
    End Function
    Function LongLongUnaryPlus() As String
        LongLongUnaryPlus = CStr(+CLngLng("9007199254740993"))
    End Function
`);
assert.strictEqual(
    arithmeticEv.callProcedure('LongLongArithmetic', []),
    '9007199254740994|9007199254740992|27021597764222979|3002399751580331|0|LongLong',
    'LongLong arithmetic preserves all 64-bit digits');
assert.strictEqual(
    arithmeticEv.callProcedure('LongLongUnaryPlus', []),
    '9007199254740993',
    'unary plus preserves LongLong');

const overflowEv = evalVBASingle(`
    Function LongLongOverflow() As LongLong
        Dim x As LongLong
        x = CLngLng("9223372036854775807")
        LongLongOverflow = x + 1
    End Function
`);
try {
    overflowEv.callProcedure('LongLongOverflow', []);
    throw new Error('LongLong overflow was not detected');
} catch (e: any) {
    assert.strictEqual(e?.type, 'VbaError', 'LongLong overflow raises a VBA error');
    assert.strictEqual(e?.number, 6, 'LongLong overflow raises Error 6');
}
console.log('✅ LongLong arithmetic preserves precision and checks range');

// XL-016: LongLong mixed with a numeric String is promoted to Double for
// display, so CStr uses Excel's 15-significant-digit scientific form.
const mixedStringEv = evalVBASingle(`
    Function LongLongStringArithmetic() As String
        Dim value As LongLong
        value = CLngLng("9007199254740993")
        LongLongStringArithmetic = CStr(value + "1") & "|" & CStr(value - "1")
    End Function
`);
assert.strictEqual(
    mixedStringEv.callProcedure('LongLongStringArithmetic', []),
    '9.00719925474099E+15|9.00719925474099E+15',
    'LongLong/String mixed arithmetic uses Excel Double CStr form');
console.log('✅ LongLong/String mixed CStr matches XL-016');

// Integer-only operators must retain an integer numeric String exactly while
// + and - above continue to use Excel's Double-promotion behavior.
const mixedIntegerStringEv = evalVBASingle(`
    Function LongLongStringIntegerOps() As String
        Dim value As LongLong, factor As String
        value = CLngLng("9007199254740993")
        factor = "2"
        LongLongStringIntegerOps = CStr(value * factor) & "|" & _
            CStr(value \\ factor) & "|" & CStr(value Mod factor) & "|" & CStr(value > factor)
    End Function
`);
assert.strictEqual(
    mixedIntegerStringEv.callProcedure('LongLongStringIntegerOps', []),
    '18014398509481986|4503599627370496|1|True',
    'LongLong/String integer operators preserve exact digits');
console.log('✅ LongLong/String integer operators preserve precision');
