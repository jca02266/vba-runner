Option Explicit

' Auto-generated test runner from vba-test-generator
' Run this Sub in Excel VBA environment to execute all tests
' Module-qualified calls: RemainingErrorTest.TestProcedures
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

    ' Execute TestCDateInvalidString
    assert.Reset
    On Error Resume Next
    RemainingErrorTest.TestCDateInvalidString assert
    On Error GoTo 0
    If Not assert.Failed Then
        testResults = testResults & "[PASS] TestCDateInvalidString" & vbCrLf
        passCount = passCount + 1
    Else
        testResults = testResults & "[FAIL] TestCDateInvalidString - " & assert.FailMessage & vbCrLf
        failCount = failCount + 1
        allPass = False
    End If

    ' Execute TestWithBlockOutsideAccess
    assert.Reset
    On Error Resume Next
    RemainingErrorTest.TestWithBlockOutsideAccess assert
    On Error GoTo 0
    If Not assert.Failed Then
        testResults = testResults & "[PASS] TestWithBlockOutsideAccess" & vbCrLf
        passCount = passCount + 1
    Else
        testResults = testResults & "[FAIL] TestWithBlockOutsideAccess - " & assert.FailMessage & vbCrLf
        failCount = failCount + 1
        allPass = False
    End If

    ' Execute TestGoToNonExistentLabel
    assert.Reset
    On Error Resume Next
    RemainingErrorTest.TestGoToNonExistentLabel assert
    On Error GoTo 0
    If Not assert.Failed Then
        testResults = testResults & "[PASS] TestGoToNonExistentLabel" & vbCrLf
        passCount = passCount + 1
    Else
        testResults = testResults & "[FAIL] TestGoToNonExistentLabel - " & assert.FailMessage & vbCrLf
        failCount = failCount + 1
        allPass = False
    End If

    ' Execute TestPrivateProcCrossModule
    assert.Reset
    On Error Resume Next
    RemainingErrorTest.TestPrivateProcCrossModule assert
    On Error GoTo 0
    If Not assert.Failed Then
        testResults = testResults & "[PASS] TestPrivateProcCrossModule" & vbCrLf
        passCount = passCount + 1
    Else
        testResults = testResults & "[FAIL] TestPrivateProcCrossModule - " & assert.FailMessage & vbCrLf
        failCount = failCount + 1
        allPass = False
    End If

    ' Execute TestCallNonCallable
    assert.Reset
    On Error Resume Next
    RemainingErrorTest.TestCallNonCallable assert
    On Error GoTo 0
    If Not assert.Failed Then
        testResults = testResults & "[PASS] TestCallNonCallable" & vbCrLf
        passCount = passCount + 1
    Else
        testResults = testResults & "[FAIL] TestCallNonCallable - " & assert.FailMessage & vbCrLf
        failCount = failCount + 1
        allPass = False
    End If

    ' Execute TestLSetNonIdentifier
    assert.Reset
    On Error Resume Next
    RemainingErrorTest.TestLSetNonIdentifier assert
    On Error GoTo 0
    If Not assert.Failed Then
        testResults = testResults & "[PASS] TestLSetNonIdentifier" & vbCrLf
        passCount = passCount + 1
    Else
        testResults = testResults & "[FAIL] TestLSetNonIdentifier - " & assert.FailMessage & vbCrLf
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
