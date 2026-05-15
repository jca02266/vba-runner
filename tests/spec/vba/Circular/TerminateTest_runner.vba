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
    On Error Resume Next
    Err.Clear
    TerminateTest.Test_BasicTerminate assert
    If Err.Number = 0 Then
        testResults = testResults & "[PASS] Test_BasicTerminate" & vbCrLf
        passCount = passCount + 1
    Else
        testResults = testResults & "[FAIL] Test_BasicTerminate" & vbCrLf
        failCount = failCount + 1
        allPass = False
    End If
    On Error GoTo 0

    ' Execute Test_MutualReferences
    On Error Resume Next
    Err.Clear
    TerminateTest.Test_MutualReferences assert
    If Err.Number = 0 Then
        testResults = testResults & "[PASS] Test_MutualReferences" & vbCrLf
        passCount = passCount + 1
    Else
        testResults = testResults & "[FAIL] Test_MutualReferences" & vbCrLf
        failCount = failCount + 1
        allPass = False
    End If
    On Error GoTo 0

    ' Execute Test_TerminateNotCalledTwice
    On Error Resume Next
    Err.Clear
    TerminateTest.Test_TerminateNotCalledTwice assert
    If Err.Number = 0 Then
        testResults = testResults & "[PASS] Test_TerminateNotCalledTwice" & vbCrLf
        passCount = passCount + 1
    Else
        testResults = testResults & "[FAIL] Test_TerminateNotCalledTwice" & vbCrLf
        failCount = failCount + 1
        allPass = False
    End If
    On Error GoTo 0

    ' Execute Test_MultipleObjects
    On Error Resume Next
    Err.Clear
    TerminateTest.Test_MultipleObjects assert
    If Err.Number = 0 Then
        testResults = testResults & "[PASS] Test_MultipleObjects" & vbCrLf
        passCount = passCount + 1
    Else
        testResults = testResults & "[FAIL] Test_MultipleObjects" & vbCrLf
        failCount = failCount + 1
        allPass = False
    End If
    On Error GoTo 0

    ' Execute Test_CircularChainCleanup
    On Error Resume Next
    Err.Clear
    TerminateTest.Test_CircularChainCleanup assert
    If Err.Number = 0 Then
        testResults = testResults & "[PASS] Test_CircularChainCleanup" & vbCrLf
        passCount = passCount + 1
    Else
        testResults = testResults & "[FAIL] Test_CircularChainCleanup" & vbCrLf
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
