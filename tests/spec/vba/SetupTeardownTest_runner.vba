Option Explicit

' Auto-generated test runner from vba-test-generator
' Run this Sub in Excel VBA environment to execute all tests
' Module-qualified calls: SetupTeardownTest.TestProcedures
' Note: SetUp/TearDown will be called before/after each test
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

    ' Execute Test_SetupInitializesCounter
    On Error Resume Next
    SetupTeardownTest.SetUp
    On Error GoTo 0
    On Error Resume Next
    Err.Clear
    SetupTeardownTest.Test_SetupInitializesCounter assert
    If Err.Number = 0 Then
        testResults = testResults & "[PASS] Test_SetupInitializesCounter" & vbCrLf
        passCount = passCount + 1
    Else
        testResults = testResults & "[FAIL] Test_SetupInitializesCounter" & vbCrLf
        failCount = failCount + 1
        allPass = False
    End If
    On Error GoTo 0
    On Error Resume Next
    SetupTeardownTest.TearDown
    On Error GoTo 0

    ' Execute Test_SetupInitializesState
    On Error Resume Next
    SetupTeardownTest.SetUp
    On Error GoTo 0
    On Error Resume Next
    Err.Clear
    SetupTeardownTest.Test_SetupInitializesState assert
    If Err.Number = 0 Then
        testResults = testResults & "[PASS] Test_SetupInitializesState" & vbCrLf
        passCount = passCount + 1
    Else
        testResults = testResults & "[FAIL] Test_SetupInitializesState" & vbCrLf
        failCount = failCount + 1
        allPass = False
    End If
    On Error GoTo 0
    On Error Resume Next
    SetupTeardownTest.TearDown
    On Error GoTo 0

    ' Execute Test_CounterIncrement
    On Error Resume Next
    SetupTeardownTest.SetUp
    On Error GoTo 0
    On Error Resume Next
    Err.Clear
    SetupTeardownTest.Test_CounterIncrement assert
    If Err.Number = 0 Then
        testResults = testResults & "[PASS] Test_CounterIncrement" & vbCrLf
        passCount = passCount + 1
    Else
        testResults = testResults & "[FAIL] Test_CounterIncrement" & vbCrLf
        failCount = failCount + 1
        allPass = False
    End If
    On Error GoTo 0
    On Error Resume Next
    SetupTeardownTest.TearDown
    On Error GoTo 0

    ' Execute Test_StringConcat
    On Error Resume Next
    SetupTeardownTest.SetUp
    On Error GoTo 0
    On Error Resume Next
    Err.Clear
    SetupTeardownTest.Test_StringConcat assert
    If Err.Number = 0 Then
        testResults = testResults & "[PASS] Test_StringConcat" & vbCrLf
        passCount = passCount + 1
    Else
        testResults = testResults & "[FAIL] Test_StringConcat" & vbCrLf
        failCount = failCount + 1
        allPass = False
    End If
    On Error GoTo 0
    On Error Resume Next
    SetupTeardownTest.TearDown
    On Error GoTo 0

    ' Execute Test_LogicalAnd
    On Error Resume Next
    SetupTeardownTest.SetUp
    On Error GoTo 0
    On Error Resume Next
    Err.Clear
    SetupTeardownTest.Test_LogicalAnd assert
    If Err.Number = 0 Then
        testResults = testResults & "[PASS] Test_LogicalAnd" & vbCrLf
        passCount = passCount + 1
    Else
        testResults = testResults & "[FAIL] Test_LogicalAnd" & vbCrLf
        failCount = failCount + 1
        allPass = False
    End If
    On Error GoTo 0
    On Error Resume Next
    SetupTeardownTest.TearDown
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
