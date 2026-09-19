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
