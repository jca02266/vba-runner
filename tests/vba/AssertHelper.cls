Option Explicit

Public Failed As Boolean
Public FailMessage As String

Sub Reset()
    Failed = False
    FailMessage = ""
End Sub

Sub Assert(actual, expected, message)
    If actual <> expected Then
        Debug.Print "[FAIL] " & message
        Debug.Print "  Expected: " & CStr(expected)
        Debug.Print "  Actual  : " & CStr(actual)
        Failed = True
        FailMessage = message
        Err.Raise vbObjectError + 1, "Assert", message
    End If
End Sub

Sub IsTrue(value, message)
    If Not CBool(value) Then
        Debug.Print "[FAIL] " & message & " (expected True)"
        Failed = True
        FailMessage = message
        Err.Raise vbObjectError + 1, "Assert", message
    End If
End Sub

Sub IsFalse(value, message)
    If CBool(value) Then
        Debug.Print "[FAIL] " & message & " (expected False)"
        Failed = True
        FailMessage = message
        Err.Raise vbObjectError + 1, "Assert", message
    End If
End Sub
