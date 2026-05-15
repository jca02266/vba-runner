Option Explicit

' Auto-generated test runner from vba-test-generator
' Run this Sub in Excel VBA environment to execute all tests
' Module-qualified calls: ArgCountTest.TestProcedures
Sub RunAllTests()
    Dim allPass As Boolean
    Dim passCount As Integer
    Dim failCount As Integer
    Dim testResults As String
    Dim assert As New AssertHelper

    allPass = True
    passCount = 0
    failCount = 0
    testResults = "=== Test Results ===" & vbCrLf & vbCrLf

    ' Execute Test_TooManyArgsToFunction
    On Error Resume Next
    Err.Clear
    ArgCountTest.Test_TooManyArgsToFunction assert
    If Err.Number = 0 Then
        testResults = testResults & "[PASS] Test_TooManyArgsToFunction" & vbCrLf
        passCount = passCount + 1
    Else
        testResults = testResults & "[FAIL] Test_TooManyArgsToFunction" & vbCrLf
        failCount = failCount + 1
        allPass = False
    End If
    On Error GoTo 0

    ' Execute Test_TooFewArgsToFunction
    On Error Resume Next
    Err.Clear
    ArgCountTest.Test_TooFewArgsToFunction assert
    If Err.Number = 0 Then
        testResults = testResults & "[PASS] Test_TooFewArgsToFunction" & vbCrLf
        passCount = passCount + 1
    Else
        testResults = testResults & "[FAIL] Test_TooFewArgsToFunction" & vbCrLf
        failCount = failCount + 1
        allPass = False
    End If
    On Error GoTo 0

    ' Execute Test_ArgsToNoArgFunction
    On Error Resume Next
    Err.Clear
    ArgCountTest.Test_ArgsToNoArgFunction assert
    If Err.Number = 0 Then
        testResults = testResults & "[PASS] Test_ArgsToNoArgFunction" & vbCrLf
        passCount = passCount + 1
    Else
        testResults = testResults & "[FAIL] Test_ArgsToNoArgFunction" & vbCrLf
        failCount = failCount + 1
        allPass = False
    End If
    On Error GoTo 0

    ' Execute Test_TooManyArgsToSub
    On Error Resume Next
    Err.Clear
    ArgCountTest.Test_TooManyArgsToSub assert
    If Err.Number = 0 Then
        testResults = testResults & "[PASS] Test_TooManyArgsToSub" & vbCrLf
        passCount = passCount + 1
    Else
        testResults = testResults & "[FAIL] Test_TooManyArgsToSub" & vbCrLf
        failCount = failCount + 1
        allPass = False
    End If
    On Error GoTo 0

    ' Execute Test_ArgsToNoArgSub
    On Error Resume Next
    Err.Clear
    ArgCountTest.Test_ArgsToNoArgSub assert
    If Err.Number = 0 Then
        testResults = testResults & "[PASS] Test_ArgsToNoArgSub" & vbCrLf
        passCount = passCount + 1
    Else
        testResults = testResults & "[FAIL] Test_ArgsToNoArgSub" & vbCrLf
        failCount = failCount + 1
        allPass = False
    End If
    On Error GoTo 0

    testResults = testResults & vbCrLf
    testResults = testResults & "=== Test Summary ===" & vbCrLf
    testResults = testResults & "Total: " & (passCount + failCount) & vbCrLf
    testResults = testResults & "Passed: " & passCount & vbCrLf
    testResults = testResults & "Failed: " & failCount & vbCrLf
    testResults = testResults & "=== Test Complete ===" & vbCrLf

    Debug.Print testResults
    MsgBox testResults, IIf(allPass, vbInformation, vbCritical), "Test Results"
End Sub
