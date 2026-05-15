Option Explicit

' Assertion class for VBA test procedures
' Fails by raising Err, with detailed output to Debug.Print

Sub Assert(actual, expected, message)
    If actual <> expected Then
        Debug.Print "[FAIL] " & message
        Debug.Print "  Expected: " & CStr(expected)
        Debug.Print "  Actual  : " & CStr(actual)
        Err.Raise vbObjectError + 1, "Assert", message
    End If
End Sub

Sub IsTrue(value, message)
    If Not CBool(value) Then
        Debug.Print "[FAIL] " & message & " (expected True)"
        Err.Raise vbObjectError + 1, "Assert", message
    End If
End Sub

Sub IsFalse(value, message)
    If CBool(value) Then
        Debug.Print "[FAIL] " & message & " (expected False)"
        Err.Raise vbObjectError + 1, "Assert", message
    End If
End Sub
