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

Public Function HexMinimum() As String
    HexMinimum = CStr(&H8000000000000000^)
End Function

Public Function HexMinusOne() As String
    HexMinusOne = CStr(&HFFFFFFFFFFFFFFFF^)
End Function

Public Function OctMinusOne() As String
    OctMinusOne = CStr(&O1777777777777777777777^)
End Function

Public Function FloatingPromotion() As String
    Dim value As Variant
    value = 9007199254740993^ + 1#
    FloatingPromotion = TypeName(value)
End Function

Public Function SinglePromotion() As String
    Dim value As Variant
    value = 9007199254740993^ + 1!
    SinglePromotion = TypeName(value)
End Function

Public Function VariantSinglePromotion() As String
    Dim value As Variant, singleValue As Single
    value = CLngLng("9007199254740993")
    singleValue = 1.5!
    VariantSinglePromotion = TypeName(value + singleValue)
End Function

Private Function VariantArgumentPromotion(ByVal value As Variant) As String
    VariantArgumentPromotion = TypeName(value + 1.5!)
End Function

Public Function VariantContainerPromotion() As String
    Dim values(0 To 0) As Variant, returned As Variant
    values(0) = CLngLng("9007199254740993")
    returned = VariantArgumentPromotion(values(0))
    VariantContainerPromotion = TypeName(values(0) + 1.5!) & ":" & returned
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
assert.equal(ev.callProcedure('HexMinimum', []), '-9223372036854775808');
assert.equal(ev.callProcedure('HexMinusOne', []), '-1');
assert.equal(ev.callProcedure('OctMinusOne', []), '-1');
assert.equal(ev.callProcedure('FloatingPromotion', []), 'Double');
assert.equal(ev.callProcedure('SinglePromotion', []), 'Double');
assert.equal(ev.callProcedure('VariantSinglePromotion', []), 'Double');
assert.equal(ev.callProcedure('VariantContainerPromotion', []), 'Double:Double');
assert.equal(ev.callProcedure('OverflowPositive', []), 6);
assert.equal(ev.callProcedure('OverflowNegative', []), 6);

console.log('[PASS] LongLong literal range and radix precision');
