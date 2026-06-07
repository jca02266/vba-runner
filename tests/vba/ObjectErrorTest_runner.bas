Option Explicit

' Auto-generated test runner from vba-test-generator
' Run this Sub in Excel VBA environment to execute all tests
' Module-qualified calls: ObjectErrorTest.TestProcedures
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

    ' Execute TestSetNonObject
    assert.Reset
    On Error Resume Next
    ObjectErrorTest.TestSetNonObject assert
    On Error GoTo 0
    If Not assert.Failed Then
        testResults = testResults & "[PASS] TestSetNonObject" & vbCrLf
        passCount = passCount + 1
    Else
        testResults = testResults & "[FAIL] TestSetNonObject - " & assert.FailMessage & vbCrLf
        failCount = failCount + 1
        allPass = False
    End If

    ' Execute TestSetString
    assert.Reset
    On Error Resume Next
    ObjectErrorTest.TestSetString assert
    On Error GoTo 0
    If Not assert.Failed Then
        testResults = testResults & "[PASS] TestSetString" & vbCrLf
        passCount = passCount + 1
    Else
        testResults = testResults & "[FAIL] TestSetString - " & assert.FailMessage & vbCrLf
        failCount = failCount + 1
        allPass = False
    End If

    ' Execute TestAccessNothingProperty
    assert.Reset
    On Error Resume Next
    ObjectErrorTest.TestAccessNothingProperty assert
    On Error GoTo 0
    If Not assert.Failed Then
        testResults = testResults & "[PASS] TestAccessNothingProperty" & vbCrLf
        passCount = passCount + 1
    Else
        testResults = testResults & "[FAIL] TestAccessNothingProperty - " & assert.FailMessage & vbCrLf
        failCount = failCount + 1
        allPass = False
    End If

    ' Execute TestCallNothingMethod
    assert.Reset
    On Error Resume Next
    ObjectErrorTest.TestCallNothingMethod assert
    On Error GoTo 0
    If Not assert.Failed Then
        testResults = testResults & "[PASS] TestCallNothingMethod" & vbCrLf
        passCount = passCount + 1
    Else
        testResults = testResults & "[FAIL] TestCallNothingMethod - " & assert.FailMessage & vbCrLf
        failCount = failCount + 1
        allPass = False
    End If

    ' Execute TestCallNonExistentMethod
    assert.Reset
    On Error Resume Next
    ObjectErrorTest.TestCallNonExistentMethod assert
    On Error GoTo 0
    If Not assert.Failed Then
        testResults = testResults & "[PASS] TestCallNonExistentMethod" & vbCrLf
        passCount = passCount + 1
    Else
        testResults = testResults & "[FAIL] TestCallNonExistentMethod - " & assert.FailMessage & vbCrLf
        failCount = failCount + 1
        allPass = False
    End If

    ' Execute TestReadNonExistentProperty
    assert.Reset
    On Error Resume Next
    ObjectErrorTest.TestReadNonExistentProperty assert
    On Error GoTo 0
    If Not assert.Failed Then
        testResults = testResults & "[PASS] TestReadNonExistentProperty" & vbCrLf
        passCount = passCount + 1
    Else
        testResults = testResults & "[FAIL] TestReadNonExistentProperty - " & assert.FailMessage & vbCrLf
        failCount = failCount + 1
        allPass = False
    End If

    ' Execute TestLateBindingNonExistentMethod
    assert.Reset
    On Error Resume Next
    ObjectErrorTest.TestLateBindingNonExistentMethod assert
    On Error GoTo 0
    If Not assert.Failed Then
        testResults = testResults & "[PASS] TestLateBindingNonExistentMethod" & vbCrLf
        passCount = passCount + 1
    Else
        testResults = testResults & "[FAIL] TestLateBindingNonExistentMethod - " & assert.FailMessage & vbCrLf
        failCount = failCount + 1
        allPass = False
    End If

    ' Execute TestLateBindingNonExistentProperty
    assert.Reset
    On Error Resume Next
    ObjectErrorTest.TestLateBindingNonExistentProperty assert
    On Error GoTo 0
    If Not assert.Failed Then
        testResults = testResults & "[PASS] TestLateBindingNonExistentProperty" & vbCrLf
        passCount = passCount + 1
    Else
        testResults = testResults & "[FAIL] TestLateBindingNonExistentProperty - " & assert.FailMessage & vbCrLf
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
