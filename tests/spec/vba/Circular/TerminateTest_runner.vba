Option Explicit

' Auto-generated test runner from vba-test-generator
' Run this Sub in Excel VBA environment to execute all tests
' Module-qualified calls: TerminateTest.TestProcedures
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

    ' Execute Test_BasicTerminate
    assert.Reset
    On Error Resume Next
    TerminateTest.Test_BasicTerminate assert
    On Error GoTo 0
    If Not assert.Failed Then
        testResults = testResults & "[PASS] Test_BasicTerminate" & vbCrLf
        passCount = passCount + 1
    Else
        testResults = testResults & "[FAIL] Test_BasicTerminate - " & assert.FailMessage & vbCrLf
        failCount = failCount + 1
        allPass = False
    End If

    ' Execute Test_MutualReferences
    assert.Reset
    On Error Resume Next
    TerminateTest.Test_MutualReferences assert
    On Error GoTo 0
    If Not assert.Failed Then
        testResults = testResults & "[PASS] Test_MutualReferences" & vbCrLf
        passCount = passCount + 1
    Else
        testResults = testResults & "[FAIL] Test_MutualReferences - " & assert.FailMessage & vbCrLf
        failCount = failCount + 1
        allPass = False
    End If

    ' Execute Test_TerminateNotCalledTwice
    assert.Reset
    On Error Resume Next
    TerminateTest.Test_TerminateNotCalledTwice assert
    On Error GoTo 0
    If Not assert.Failed Then
        testResults = testResults & "[PASS] Test_TerminateNotCalledTwice" & vbCrLf
        passCount = passCount + 1
    Else
        testResults = testResults & "[FAIL] Test_TerminateNotCalledTwice - " & assert.FailMessage & vbCrLf
        failCount = failCount + 1
        allPass = False
    End If

    ' Execute Test_MultipleObjects
    assert.Reset
    On Error Resume Next
    TerminateTest.Test_MultipleObjects assert
    On Error GoTo 0
    If Not assert.Failed Then
        testResults = testResults & "[PASS] Test_MultipleObjects" & vbCrLf
        passCount = passCount + 1
    Else
        testResults = testResults & "[FAIL] Test_MultipleObjects - " & assert.FailMessage & vbCrLf
        failCount = failCount + 1
        allPass = False
    End If

    ' Execute Test_CircularChainCleanup
    assert.Reset
    On Error Resume Next
    TerminateTest.Test_CircularChainCleanup assert
    On Error GoTo 0
    If Not assert.Failed Then
        testResults = testResults & "[PASS] Test_CircularChainCleanup" & vbCrLf
        passCount = passCount + 1
    Else
        testResults = testResults & "[FAIL] Test_CircularChainCleanup - " & assert.FailMessage & vbCrLf
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
