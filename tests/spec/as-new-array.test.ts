/** `As New` クラス型配列の要素ごとの遅延インスタンス化。 */
import { evalVBAModules, assert } from '../../test-libs/test-runner';

const output: string[] = [];
const ticket = `
Option Explicit
Private mCode As String
Public Property Get Code() As String
    Code = mCode
End Property
Public Property Let Code(ByVal value As String)
    mCode = value
End Property
`;
const main = `
Sub Test()
    Dim tickets(0 To 1) As New Ticket
    tickets(0).Code = "X"
    tickets(1).Code = "Y"
    Debug.Print tickets(0).Code & ":" & tickets(1).Code
End Sub
`;

const evaluator = evalVBAModules(
    [
        { name: 'Ticket', code: ticket, parseAsClass: 'Ticket' },
        { name: 'Main', code: main },
    ],
    { onPrint: (line) => output.push(line) },
);
evaluator.callProcedure('Test', []);
assert.deepStrictEqual(output, ['X:Y'], 'As New array elements retain independent instances');
console.log('[PASS] As New クラス型配列: 要素ごとの遅延インスタンス化');

const expandedOutput: string[] = [];
const holder = `
Option Explicit
Public Items(0 To 1) As New Ticket
`;
const expandedMain = `
Sub Test()
    Dim h As New Holder
    h.Items(0).Code = "F"
    Dim items() As New Ticket
    ReDim items(0 To 1)
    items(0).Code = "R"
    ReDim Preserve items(0 To 2)
    items(2).Code = "N"
    Erase items
    ReDim items(0 To 0)
    items(0).Code = "E"
    Debug.Print h.Items(0).Code & ":" & items(0).Code
End Sub
`;
const expanded = evalVBAModules(
    [
        { name: 'Ticket', code: ticket, parseAsClass: 'Ticket' },
        { name: 'Holder', code: holder, parseAsClass: 'Holder' },
        { name: 'Main', code: expandedMain },
    ],
    { onPrint: (line) => expandedOutput.push(line) },
);
expanded.callProcedure('Test', []);
assert.deepStrictEqual(expandedOutput, ['F:E'], 'class fields, ReDim Preserve, and Erase retain As New behavior');
console.log('[PASS] As New クラス型配列: フィールド・ReDim・Erase横展開');
