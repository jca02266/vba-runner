import { evalVBASingle, assert } from '../../test-libs/test-runner';

const ev = evalVBASingle(String.raw`
Function ReadComArray() As Long
    Dim source As Object, target() As Long
    Set source = CreateObject("Test.Array")
    target = source.Values
    ReadComArray = target(0)
End Function

Function ReadComObjectArray() As Long
    Dim source As Object, target(0 To 0) As Object
    Dim errNo As Long
    Set source = CreateObject("Test.ObjectArray")
    On Error Resume Next
    target = source.Values
    errNo = Err.Number
    ReadComObjectArray = errNo
End Function
`);

const longValues: any[] = [7];
(longValues as any).__vbaElementType__ = 'long';
ev.registerComObject(() => ({
    __progId__: 'Test.Array',
    Values: longValues,
}) as any);

assert.strictEqual(ev.callProcedure('ReadComArray', []), 7,
    'COM mock SAFEARRAY getter can initialize a typed array');
console.log('[PASS] COM mock array getter boundary');

let objectValues: any[] = [{ __vbaClass__: true, __className__: 'Payload' }];
const objectEvaluator = evalVBASingle(String.raw`
Function ReadComObjectArray() As Long
    Dim source As Object, target(0 To 0) As Object
    Dim errNo As Long
    Set source = CreateObject("Test.ObjectArray")
    On Error Resume Next
    target = source.Values
    errNo = Err.Number
    ReadComObjectArray = errNo
End Function
`);
objectEvaluator.registerComObject(() => ({
    __progId__: 'Test.ObjectArray',
    Values: objectValues,
}) as any);

assert.strictEqual(objectEvaluator.callProcedure('ReadComObjectArray', []), 0,
    'COM getter Object() accepts an object reference element');
objectValues = [42];
assert.strictEqual(objectEvaluator.callProcedure('ReadComObjectArray', []), 13,
    'COM getter Object() rejects a numeric element');
objectValues = ['text'];
assert.strictEqual(objectEvaluator.callProcedure('ReadComObjectArray', []), 13,
    'COM getter Object() rejects a string element');
objectValues = [{ __vbaClass__: true, __className__: 'Payload' }, 42];
assert.strictEqual(objectEvaluator.callProcedure('ReadComObjectArray', []), 13,
    'COM getter Object() rejects mixed element categories');
console.log('[PASS] COM mock Object() element-category boundary');

const variantEvaluator = evalVBASingle(String.raw`
Function ReadComVariantArray() As Long
    Dim source As Object, target() As Variant
    Dim errNo As Long
    Set source = CreateObject("Test.Range")
    On Error Resume Next
    target = source.Values
    errNo = Err.Number
    If errNo = 0 Then
        ReadComVariantArray = UBound(target, 2)
    Else
        ReadComVariantArray = errNo
    End If
End Function
`);
const rangeValues: any[] = [[1, 2, 3], [4, 5, 6]];
(rangeValues as any).__vbaArrayReturn__ = true;
(rangeValues as any).__vbaDimensions__ = [
    { lower: 1, upper: 2 },
    { lower: 1, upper: 3 },
];
variantEvaluator.registerComObject(() => ({
    __progId__: 'Test.Range',
    Values: rangeValues,
}) as any);
assert.strictEqual(variantEvaluator.callProcedure('ReadComVariantArray', []), 3,
    'Range-like opaque 2D Variant array preserves its second-dimension bound');
console.log('[PASS] COM mock Range-like Variant() shape boundary');

const inferredEvaluator = evalVBASingle(String.raw`
Function ReadComInferredArray() As Long
    Dim source As Object, target() As Variant
    Set source = CreateObject("Test.InferredRange")
    target = source.Values
    ReadComInferredArray = UBound(target, 2)
End Function
`);
const inferredValues: any[] = [[1, 2, 3], [4, 5, 6]];
inferredEvaluator.registerComObject(() => ({
    __progId__: 'Test.InferredRange',
    Values: inferredValues,
}) as any);
assert.strictEqual(inferredEvaluator.callProcedure('ReadComInferredArray', []), 2,
    'opaque nested arrays infer rank when the host omits SAFEARRAY bounds');
console.log('[PASS] COM mock inferred Variant() rank boundary');
