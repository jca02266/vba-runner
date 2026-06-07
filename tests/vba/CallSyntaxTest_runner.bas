Option Explicit

' Auto-generated test runner from vba-test-generator
' Run this Sub in Excel VBA environment to execute all tests
' Module-qualified calls: CallSyntaxTest.TestProcedures
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

    ' Execute TestSubCallStatement
    assert.Reset
    On Error Resume Next
    CallSyntaxTest.TestSubCallStatement assert
    On Error GoTo 0
    If Not assert.Failed Then
        testResults = testResults & "[PASS] TestSubCallStatement" & vbCrLf
        passCount = passCount + 1
    Else
        testResults = testResults & "[FAIL] TestSubCallStatement - " & assert.FailMessage & vbCrLf
        failCount = failCount + 1
        allPass = False
    End If

    ' Execute TestSubCallWithSpaceParen
    assert.Reset
    On Error Resume Next
    CallSyntaxTest.TestSubCallWithSpaceParen assert
    On Error GoTo 0
    If Not assert.Failed Then
        testResults = testResults & "[PASS] TestSubCallWithSpaceParen" & vbCrLf
        passCount = passCount + 1
    Else
        testResults = testResults & "[FAIL] TestSubCallWithSpaceParen - " & assert.FailMessage & vbCrLf
        failCount = failCount + 1
        allPass = False
    End If

    ' Execute TestFuncCallExpression
    assert.Reset
    On Error Resume Next
    CallSyntaxTest.TestFuncCallExpression assert
    On Error GoTo 0
    If Not assert.Failed Then
        testResults = testResults & "[PASS] TestFuncCallExpression" & vbCrLf
        passCount = passCount + 1
    Else
        testResults = testResults & "[FAIL] TestFuncCallExpression - " & assert.FailMessage & vbCrLf
        failCount = failCount + 1
        allPass = False
    End If

    ' Execute TestFuncWithArgExpression
    assert.Reset
    On Error Resume Next
    CallSyntaxTest.TestFuncWithArgExpression assert
    On Error GoTo 0
    If Not assert.Failed Then
        testResults = testResults & "[PASS] TestFuncWithArgExpression" & vbCrLf
        passCount = passCount + 1
    Else
        testResults = testResults & "[FAIL] TestFuncWithArgExpression - " & assert.FailMessage & vbCrLf
        failCount = failCount + 1
        allPass = False
    End If

    ' Execute TestCallKeywordNoParens
    assert.Reset
    On Error Resume Next
    CallSyntaxTest.TestCallKeywordNoParens assert
    On Error GoTo 0
    If Not assert.Failed Then
        testResults = testResults & "[PASS] TestCallKeywordNoParens" & vbCrLf
        passCount = passCount + 1
    Else
        testResults = testResults & "[FAIL] TestCallKeywordNoParens - " & assert.FailMessage & vbCrLf
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
