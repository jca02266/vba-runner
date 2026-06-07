Option Explicit

' Auto-generated test runner from vba-test-generator
' Run this Sub in Excel VBA environment to execute all tests
' Module-qualified calls: ProcNotFoundTest.TestProcedures
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

    ' Execute TestUndefinedFuncWithArgs
    assert.Reset
    On Error Resume Next
    ProcNotFoundTest.TestUndefinedFuncWithArgs assert
    On Error GoTo 0
    If Not assert.Failed Then
        testResults = testResults & "[PASS] TestUndefinedFuncWithArgs" & vbCrLf
        passCount = passCount + 1
    Else
        testResults = testResults & "[FAIL] TestUndefinedFuncWithArgs - " & assert.FailMessage & vbCrLf
        failCount = failCount + 1
        allPass = False
    End If

    ' Execute TestUndefinedSubWithArgs
    assert.Reset
    On Error Resume Next
    ProcNotFoundTest.TestUndefinedSubWithArgs assert
    On Error GoTo 0
    If Not assert.Failed Then
        testResults = testResults & "[PASS] TestUndefinedSubWithArgs" & vbCrLf
        passCount = passCount + 1
    Else
        testResults = testResults & "[FAIL] TestUndefinedSubWithArgs - " & assert.FailMessage & vbCrLf
        failCount = failCount + 1
        allPass = False
    End If

    ' Execute TestNewUnknownClass
    assert.Reset
    On Error Resume Next
    ProcNotFoundTest.TestNewUnknownClass assert
    On Error GoTo 0
    If Not assert.Failed Then
        testResults = testResults & "[PASS] TestNewUnknownClass" & vbCrLf
        passCount = passCount + 1
    Else
        testResults = testResults & "[FAIL] TestNewUnknownClass - " & assert.FailMessage & vbCrLf
        failCount = failCount + 1
        allPass = False
    End If

    ' Execute TestCreateObjectUnknown
    assert.Reset
    On Error Resume Next
    ProcNotFoundTest.TestCreateObjectUnknown assert
    On Error GoTo 0
    If Not assert.Failed Then
        testResults = testResults & "[PASS] TestCreateObjectUnknown" & vbCrLf
        passCount = passCount + 1
    Else
        testResults = testResults & "[FAIL] TestCreateObjectUnknown - " & assert.FailMessage & vbCrLf
        failCount = failCount + 1
        allPass = False
    End If

    testResults = testResults & vbCrLf
    testResults = testResults & "=== Test Summary ===" & vbCrLf
    testResults = testResults & "Total: " & (passCount + failCount) & vbCrLf
    testResults = testResults & "Passed: " & passCount & vbCrLf
    testResults = testResults & "Failed: " & failCount & vbCrLf
    testResults = testResults & "=== Test Complete ===" & vbCrLf

    Debug.Print testResults
    MsgBox testResults, IIf(allPass, vbInformation, vbCritical), "Test Results"
End Sub
