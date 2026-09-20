import { evalVBASingle, evalVBAModules, assert } from '../../test-libs/test-runner';

// UDT member bounds and fixed-length strings use the same constant-expression
// contract as ordinary array declarations.
{
    const source = String.raw`Option Explicit
Const UDT_LAST_INDEX As Integer = 1
Type IndexedRecord
    Values(UDT_LAST_INDEX) As Integer
End Type
Function RunCase() As String
    Dim record As IndexedRecord
    record.Values(0) = 7
    record.Values(1) = 9
    RunCase = CStr(record.Values(0)) & "," & CStr(record.Values(1))
End Function`;

    const result = evalVBASingle(source).callProcedure('RunCase', []);
    assert.strictEqual(result, '7,9',
        'UDT member upper bounds resolve an unqualified Const expression');
    console.log('[PASS] UDT member Const upper bound');
}

{
    const source = String.raw`Option Explicit
Const UDT_LAST_INDEX As Integer = 1
Const UDT_STRING_LENGTH As Integer = 3
Type IndexedRecord
    Values(0 To UDT_LAST_INDEX, 0 To 1) As Integer
    Label As String * UDT_STRING_LENGTH
End Type
Function RunCase() As String
    Dim record As IndexedRecord
    record.Values(1, 1) = 9
    record.Label = "ABCDE"
    RunCase = CStr(record.Values(1, 1)) & "|" & CStr(Len(record.Label)) & "|" & record.Label
End Function`;

    const result = evalVBASingle(source).callProcedure('RunCase', []);
    assert.strictEqual(result, '9|3|ABC',
        'UDT member array and fixed-length string bounds resolve Const expressions');
    console.log('[PASS] UDT member Const bounds');
}

{
    const evaluator = evalVBAModules([
        {
            name: 'RecordHolder',
            parseAsClass: 'RecordHolder',
            code: String.raw`Const LAST_INDEX As Integer = 1
Type InnerRecord
    Values(LAST_INDEX) As Integer
End Type
Public Function Build() As String
    Dim record As InnerRecord
    record.Values(1) = 11
    Build = CStr(record.Values(1))
End Function`,
        },
        {
            name: 'Module1',
            code: String.raw`Function RunCase() As String
    Dim holder As New RecordHolder
    RunCase = holder.Build()
End Function`,
        },
    ]);
    assert.strictEqual(evaluator.callProcedure('RunCase', []), '11',
        'Class-local UDT member bounds resolve Const expressions');
    console.log('[PASS] Class-local UDT member Const bound');
}

{
    const source = String.raw`Option Explicit
Type IndexedRecord
    Values(NOT_A_CONST) As Integer
End Type
Function RunCase() As String
    RunCase = "unreachable"
End Function`;
    assert.throwsMatch(
        () => evalVBASingle(source),
        /Constant expression required|Invalid array bound/,
        'UDT member bounds reject non-constant identifiers',
    );
    console.log('[PASS] UDT member non-Const bound rejected');
}

{
    const source = String.raw`Option Explicit
Option Base 1
Const UDT_LAST_INDEX As Integer = 2
Type IndexedRecord
    Values(UDT_LAST_INDEX) As Integer
End Type
Function RunCase() As String
    Dim record As IndexedRecord
    record.Values(1) = 7
    record.Values(2) = 9
    RunCase = CStr(LBound(record.Values)) & ":" & CStr(UBound(record.Values)) & _
        ":" & CStr(record.Values(1)) & ":" & CStr(record.Values(2))
End Function`;

    const result = evalVBASingle(source).callProcedure('RunCase', []);
    assert.strictEqual(result, '1:2:7:9',
        'UDT member omitted lower bounds use the containing module Option Base');
    console.log('[PASS] UDT member Option Base bound');
}

{
    const source = String.raw`Option Explicit
Const UDT_LAST_INDEX As Double = 1.5
Type IndexedRecord
    Values(UDT_LAST_INDEX) As Integer
End Type
Function RunCase() As String
    Dim record As IndexedRecord
    record.Values(2) = 9
    RunCase = CStr(UBound(record.Values)) & ":" & CStr(record.Values(2))
End Function`;

    const result = evalVBASingle(source).callProcedure('RunCase', []);
    assert.strictEqual(result, '2:9',
        'UDT member bounds apply VBA Long Let-coercion with banker rounding');
    console.log('[PASS] UDT member Long-coercible bound');
}

{
    const evaluator = evalVBAModules([
        {
            name: 'Bounds',
            code: String.raw`Public Const LAST_INDEX As Integer = 2`,
        },
        {
            name: 'Records',
            code: String.raw`Option Explicit
Type IndexedRecord
    Values(Bounds.LAST_INDEX) As Integer
End Type
Function RunCase() As String
    Dim record As IndexedRecord
    record.Values(2) = 9
    RunCase = CStr(UBound(record.Values)) & ":" & CStr(record.Values(2))
End Function`,
        },
    ]);
    assert.strictEqual(evaluator.callProcedure('RunCase', []), '2:9',
        'UDT member bounds resolve a qualified Public Const');
    console.log('[PASS] UDT member qualified Public Const bound');
}

{
    assert.throwsMatch(
        () => evalVBAModules([
            {
                name: 'PrivateBounds',
                code: String.raw`Private Const LAST_INDEX As Integer = 2`,
            },
            {
                name: 'Records',
                code: String.raw`Option Explicit
Type IndexedRecord
    Values(PrivateBounds.LAST_INDEX) As Integer
End Type`,
            },
        ]),
        /Private constant|Constant expression required/,
        'UDT member bounds reject a qualified Private Const from another module',
    );
    console.log('[PASS] UDT member qualified Private Const rejected');
}

{
    const evaluator = evalVBAModules([
        {
            name: 'Records',
            code: String.raw`Option Explicit
Private Const LAST_INDEX As Integer = 2
Type IndexedRecord
    Values(LAST_INDEX) As Integer
End Type
Function RunCase() As Long
    Dim record As IndexedRecord
    RunCase = UBound(record.Values)
End Function`,
        },
        {
            name: 'Other',
            code: String.raw`Public Const LAST_INDEX As Integer = 4`,
        },
    ]);
    assert.strictEqual(evaluator.callProcedure('RunCase', []), 2,
        'An unqualified UDT bound resolves in its owning module before global constants');
    console.log('[PASS] UDT member owner-module Private Const bound');
}

{
    const source = String.raw`Option Explicit
Enum RecordLimit
    LastRecord = 3
End Enum
Type IndexedRecord
    Converted(CLng(1.5)) As Integer
    Enumerated(RecordLimit.LastRecord) As Integer
End Type
Function RunCase() As String
    Dim record As IndexedRecord
    RunCase = CStr(UBound(record.Converted)) & ":" & CStr(UBound(record.Enumerated))
End Function`;
    assert.strictEqual(evalVBASingle(source).callProcedure('RunCase', []), '2:3',
        'UDT bounds accept the intrinsic calls and Enum members allowed in constant expressions');
    console.log('[PASS] UDT member intrinsic and Enum constant expressions');
}
