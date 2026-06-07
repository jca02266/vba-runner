Option Explicit

' Auto-generated test runner from vba-test-generator
' Run this Sub in Excel VBA environment to execute all tests
' Module-qualified calls: MiscErrorTest.TestProcedures
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

    ' Execute TestAssignToConst
    assert.Reset
    On Error Resume Next
    MiscErrorTest.TestAssignToConst assert
    On Error GoTo 0
    If Not assert.Failed Then
        testResults = testResults & "[PASS] TestAssignToConst" & vbCrLf
        passCount = passCount + 1
    Else
        testResults = testResults & "[FAIL] TestAssignToConst - " & assert.FailMessage & vbCrLf
        failCount = failCount + 1
        allPass = False
    End If

    ' Execute TestOnGotoOutOfRange
    assert.Reset
    On Error Resume Next
    MiscErrorTest.TestOnGotoOutOfRange assert
    On Error GoTo 0
    If Not assert.Failed Then
        testResults = testResults & "[PASS] TestOnGotoOutOfRange" & vbCrLf
        passCount = passCount + 1
    Else
        testResults = testResults & "[FAIL] TestOnGotoOutOfRange - " & assert.FailMessage & vbCrLf
        failCount = failCount + 1
        allPass = False
    End If

    ' Execute TestInvalidArrayBounds
    assert.Reset
    On Error Resume Next
    MiscErrorTest.TestInvalidArrayBounds assert
    On Error GoTo 0
    If Not assert.Failed Then
        testResults = testResults & "[PASS] TestInvalidArrayBounds" & vbCrLf
        passCount = passCount + 1
    Else
        testResults = testResults & "[FAIL] TestInvalidArrayBounds - " & assert.FailMessage & vbCrLf
        failCount = failCount + 1
        allPass = False
    End If

    ' Execute TestBangOnCollection
    assert.Reset
    On Error Resume Next
    MiscErrorTest.TestBangOnCollection assert
    On Error GoTo 0
    If Not assert.Failed Then
        testResults = testResults & "[PASS] TestBangOnCollection" & vbCrLf
        passCount = passCount + 1
    Else
        testResults = testResults & "[FAIL] TestBangOnCollection - " & assert.FailMessage & vbCrLf
        failCount = failCount + 1
        allPass = False
    End If

    ' Execute TestDefaultPropertyMissing
    assert.Reset
    On Error Resume Next
    MiscErrorTest.TestDefaultPropertyMissing assert
    On Error GoTo 0
    If Not assert.Failed Then
        testResults = testResults & "[PASS] TestDefaultPropertyMissing" & vbCrLf
        passCount = passCount + 1
    Else
        testResults = testResults & "[FAIL] TestDefaultPropertyMissing - " & assert.FailMessage & vbCrLf
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
