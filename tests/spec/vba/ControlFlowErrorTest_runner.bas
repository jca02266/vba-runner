Option Explicit

' Auto-generated test runner from vba-test-generator
' Run this Sub in Excel VBA environment to execute all tests
' Module-qualified calls: ControlFlowErrorTest.TestProcedures
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

    ' Execute TestReturnWithoutGoSub
    assert.Reset
    On Error Resume Next
    ControlFlowErrorTest.TestReturnWithoutGoSub assert
    On Error GoTo 0
    If Not assert.Failed Then
        testResults = testResults & "[PASS] TestReturnWithoutGoSub" & vbCrLf
        passCount = passCount + 1
    Else
        testResults = testResults & "[FAIL] TestReturnWithoutGoSub - " & assert.FailMessage & vbCrLf
        failCount = failCount + 1
        allPass = False
    End If

    ' Execute TestForEachNonCollection
    assert.Reset
    On Error Resume Next
    ControlFlowErrorTest.TestForEachNonCollection assert
    On Error GoTo 0
    If Not assert.Failed Then
        testResults = testResults & "[PASS] TestForEachNonCollection" & vbCrLf
        passCount = passCount + 1
    Else
        testResults = testResults & "[FAIL] TestForEachNonCollection - " & assert.FailMessage & vbCrLf
        failCount = failCount + 1
        allPass = False
    End If

    ' Execute TestWriteToClosedFile
    assert.Reset
    On Error Resume Next
    ControlFlowErrorTest.TestWriteToClosedFile assert
    On Error GoTo 0
    If Not assert.Failed Then
        testResults = testResults & "[PASS] TestWriteToClosedFile" & vbCrLf
        passCount = passCount + 1
    Else
        testResults = testResults & "[FAIL] TestWriteToClosedFile - " & assert.FailMessage & vbCrLf
        failCount = failCount + 1
        allPass = False
    End If

    ' Execute TestInputFromClosedFile
    assert.Reset
    On Error Resume Next
    ControlFlowErrorTest.TestInputFromClosedFile assert
    On Error GoTo 0
    If Not assert.Failed Then
        testResults = testResults & "[PASS] TestInputFromClosedFile" & vbCrLf
        passCount = passCount + 1
    Else
        testResults = testResults & "[FAIL] TestInputFromClosedFile - " & assert.FailMessage & vbCrLf
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
