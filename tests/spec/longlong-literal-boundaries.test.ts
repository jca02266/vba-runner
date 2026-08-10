import { strict as assert } from 'node:assert';
import { evalVBASingle } from '../../test-libs/test-runner';

const ev = evalVBASingle(String.raw`
Option Explicit

Public Function ExactHex() As String
    ExactHex = CStr(&H20000000000001^)
End Function

Public Function ExactOct() As String
    ExactOct = CStr(&O400000000000000000001^)
End Function

Public Function Minimum() As String
    Minimum = CStr(-9223372036854775808^)
End Function

Public Function OverflowPositive() As Long
    On Error Resume Next
    Dim value As Variant
    value = 9223372036854775808^
    OverflowPositive = Err.Number
End Function

Public Function OverflowNegative() As Long
    On Error Resume Next
    Dim value As Variant
    value = -9223372036854775809^
    OverflowNegative = Err.Number
End Function
`);

assert.equal(ev.callProcedure('ExactHex', []), '9007199254740993');
assert.equal(ev.callProcedure('ExactOct', []), '4611686018427387905');
assert.equal(ev.callProcedure('Minimum', []), '-9223372036854775808');
assert.equal(ev.callProcedure('OverflowPositive', []), 6);
assert.equal(ev.callProcedure('OverflowNegative', []), 6);

console.log('[PASS] LongLong literal range and radix precision');
