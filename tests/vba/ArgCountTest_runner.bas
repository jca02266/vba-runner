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
    assert.Reset
    On Error Resume Next
    ArgCountTest.Test_TooManyArgsToFunction assert
    On Error GoTo 0
    If Not assert.Failed Then
        testResults = testResults & "[PASS] Test_TooManyArgsToFunction" & vbCrLf
        passCount = passCount + 1
    Else
        testResults = testResults & "[FAIL] Test_TooManyArgsToFunction - " & assert.FailMessage & vbCrLf
        failCount = failCount + 1
        allPass = False
    End If

    ' Execute Test_TooFewArgsToFunction
    assert.Reset
    On Error Resume Next
    ArgCountTest.Test_TooFewArgsToFunction assert
    On Error GoTo 0
    If Not assert.Failed Then
        testResults = testResults & "[PASS] Test_TooFewArgsToFunction" & vbCrLf
        passCount = passCount + 1
    Else
        testResults = testResults & "[FAIL] Test_TooFewArgsToFunction - " & assert.FailMessage & vbCrLf
        failCount = failCount + 1
        allPass = False
    End If

    ' Execute Test_ArgsToNoArgFunction
    assert.Reset
    On Error Resume Next
    ArgCountTest.Test_ArgsToNoArgFunction assert
    On Error GoTo 0
    If Not assert.Failed Then
        testResults = testResults & "[PASS] Test_ArgsToNoArgFunction" & vbCrLf
        passCount = passCount + 1
    Else
        testResults = testResults & "[FAIL] Test_ArgsToNoArgFunction - " & assert.FailMessage & vbCrLf
        failCount = failCount + 1
        allPass = False
    End If

    ' Execute Test_TooManyArgsToSub
    assert.Reset
    On Error Resume Next
    ArgCountTest.Test_TooManyArgsToSub assert
    On Error GoTo 0
    If Not assert.Failed Then
        testResults = testResults & "[PASS] Test_TooManyArgsToSub" & vbCrLf
        passCount = passCount + 1
    Else
        testResults = testResults & "[FAIL] Test_TooManyArgsToSub - " & assert.FailMessage & vbCrLf
        failCount = failCount + 1
        allPass = False
    End If

    ' Execute Test_ArgsToNoArgSub
    assert.Reset
    On Error Resume Next
    ArgCountTest.Test_ArgsToNoArgSub assert
    On Error GoTo 0
    If Not assert.Failed Then
        testResults = testResults & "[PASS] Test_ArgsToNoArgSub" & vbCrLf
        passCount = passCount + 1
    Else
        testResults = testResults & "[FAIL] Test_ArgsToNoArgSub - " & assert.FailMessage & vbCrLf
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
