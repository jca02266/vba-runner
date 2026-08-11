/** GoSub/ReturnとOn Error復帰境界。 */
import { evalVBAModules, evalVBASingle, assert } from '../../test-libs/test-runner';

const ev = evalVBASingle(String.raw`
Function ProbeNormal() As String
    Dim value As Long
    value = 1
    GoSub Worker
    ProbeNormal = CStr(value)
    Exit Function
Worker:
    value = value + 2
    Return
End Function

Function ProbeError() As String
    On Error GoTo Handler
    GoSub Worker
    ProbeError = "bad"
    Exit Function
Worker:
    Err.Raise 5
    Return
Handler:
    ProbeError = "H" & CStr(Err.Number)
End Function

Function ProbeResume() As String
    Dim value As Long
    On Error Resume Next
    GoSub Worker
    value = value + 1
    ProbeResume = CStr(Err.Number) & ":" & CStr(value)
    Exit Function
Worker:
    Err.Raise 6
    Return
End Function

Function ProbeCallBoundary() As String
    On Error GoTo Handler
    GoSub Worker
    ProbeCallBoundary = "bad"
    Exit Function
Worker:
    Call Helper
    Return
Handler:
    ProbeCallBoundary = "H" & CStr(Err.Number)
End Function

Function ProbeResumeLabel() As String
    On Error GoTo Handler
    GoSub Worker
    ProbeResumeLabel = "bad"
    Exit Function
Worker:
    Err.Raise 5
    Return
Handler:
    Resume Recover
Recover:
    ProbeResumeLabel = "R" & CStr(Err.Number)
End Function

Function ProbeResumeNextHandler() As String
    On Error GoTo Handler
    GoSub Worker
    ProbeResumeNextHandler = "after"
    Exit Function
Worker:
    Err.Raise 5
    Return
Handler:
    Resume Next
End Function

Sub Helper()
    Err.Raise 5
End Sub
`);

assert.strictEqual(ev.callProcedure('ProbeNormal', []), '3',
    'GoSub and Return resume the caller at the next statement');
assert.strictEqual(ev.callProcedure('ProbeError', []), 'H5',
    'GoSub errors reach the caller On Error handler');
assert.strictEqual(ev.callProcedure('ProbeResume', []), '6:1',
    'Resume Next handles an error raised inside a GoSub body');
assert.strictEqual(ev.callProcedure('ProbeCallBoundary', []), 'H5',
    'An error from a procedure called inside GoSub reaches the owning handler');
assert.strictEqual(ev.callProcedure('ProbeResumeLabel', []), 'R0',
    'A handler Resume label exits the GoSub error path without re-entry');
assert.strictEqual(ev.callProcedure('ProbeResumeNextHandler', []), 'after',
    'A handler Resume Next returns from the GoSub error statement');
console.log('[PASS] GoSub error resumption matrix');

const propertyEv = evalVBAModules([
    {
        name: 'Thing',
        parseAsClass: 'Thing',
        code: String.raw`
Option Explicit
Private mValue As Long

Public Property Get Value() As Long
    On Error GoTo Handler
    GoSub Worker
    Value = 99
    Exit Property
Worker:
    Err.Raise 5
    Return
Handler:
    Value = 42
End Property

Public Property Let Value(ByVal incoming As Long)
    On Error GoTo Handler
    GoSub Worker
    mValue = incoming
    Exit Property
Worker:
    Err.Raise 5
    Return
Handler:
    mValue = 42
End Property
`,
    },
    {
        name: 'Module1',
        code: String.raw`
Public Function ProbeProperty() As Long
    Dim t As New Thing
    ProbeProperty = t.Value
End Function

Public Function ProbePropertyLet() As Long
    Dim t As New Thing
    t.Value = 7
    ProbePropertyLet = t.Value
End Function
`,
    },
]);

assert.strictEqual(propertyEv.callProcedure('ProbeProperty', []), 42,
    'GoSub errors terminate a Property Get after its owning handler');
console.log('[PASS] GoSub Property Get error resumption');
assert.strictEqual(propertyEv.callProcedure('ProbePropertyLet', []), 42,
    'GoSub errors terminate a Property Let after its owning handler');
console.log('[PASS] GoSub Property Let error resumption');

const propertySetEv = evalVBAModules([
    {
        name: 'Payload',
        parseAsClass: 'Payload',
        code: String.raw`Public Name As String
`,
    },
    {
        name: 'Holder',
        parseAsClass: 'Holder',
        code: String.raw`
Private mItem As Object
Private mSeen As Long

Public Property Set Item(ByVal value As Object)
    On Error GoTo Handler
    GoSub Worker
    Set mItem = value
    Exit Property
Worker:
    Err.Raise 5
    Return
Handler:
    mSeen = 42
End Property

Public Property Get Seen() As Long
    Seen = mSeen
End Property
`,
    },
    {
        name: 'Module1',
        code: String.raw`
Public Function ProbePropertySet() As Long
    Dim h As New Holder, p As New Payload
    Set h.Item = p
    ProbePropertySet = h.Seen
End Function
`,
    },
]);

assert.strictEqual(propertySetEv.callProcedure('ProbePropertySet', []), 42,
    'GoSub errors terminate a Property Set after its owning handler');
console.log('[PASS] GoSub Property Set error resumption');

const dynamicEv = evalVBAModules([
    {
        name: 'WorkerObject',
        parseAsClass: 'WorkerObject',
        code: String.raw`
Public Function Boom() As Long
    Err.Raise 5
End Function
`,
    },
    {
        name: 'Module1',
        code: String.raw`
Public Function ProbeDynamicCall() As String
    Dim worker As New WorkerObject
    On Error GoTo Handler
    GoSub Worker
    ProbeDynamicCall = "bad"
    Exit Function
Worker:
    CallByName worker, "Boom", VbMethod
    Return
Handler:
    ProbeDynamicCall = "H" & CStr(Err.Number)
End Function
`,
    },
]);

assert.strictEqual(dynamicEv.callProcedure('ProbeDynamicCall', []), 'H5',
    'A dynamic call from GoSub reaches the owning handler');
console.log('[PASS] GoSub dynamic-call error resumption');
