Option Explicit

' Auto-generated test runner from vba-test-generator
' Run this Sub in Excel VBA environment to execute all tests
' Module-qualified calls: CollectionErrorTest.TestProcedures
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

    ' Execute TestDuplicateKeyError
    assert.Reset
    On Error Resume Next
    CollectionErrorTest.TestDuplicateKeyError assert
    On Error GoTo 0
    If Not assert.Failed Then
        testResults = testResults & "[PASS] TestDuplicateKeyError" & vbCrLf
        passCount = passCount + 1
    Else
        testResults = testResults & "[FAIL] TestDuplicateKeyError - " & assert.FailMessage & vbCrLf
        failCount = failCount + 1
        allPass = False
    End If

    ' Execute TestNumericIndexOutOfRange
    assert.Reset
    On Error Resume Next
    CollectionErrorTest.TestNumericIndexOutOfRange assert
    On Error GoTo 0
    If Not assert.Failed Then
        testResults = testResults & "[PASS] TestNumericIndexOutOfRange" & vbCrLf
        passCount = passCount + 1
    Else
        testResults = testResults & "[FAIL] TestNumericIndexOutOfRange - " & assert.FailMessage & vbCrLf
        failCount = failCount + 1
        allPass = False
    End If

    ' Execute TestRemoveOutOfRange
    assert.Reset
    On Error Resume Next
    CollectionErrorTest.TestRemoveOutOfRange assert
    On Error GoTo 0
    If Not assert.Failed Then
        testResults = testResults & "[PASS] TestRemoveOutOfRange" & vbCrLf
        passCount = passCount + 1
    Else
        testResults = testResults & "[FAIL] TestRemoveOutOfRange - " & assert.FailMessage & vbCrLf
        failCount = failCount + 1
        allPass = False
    End If

    ' Execute TestKeyNotFound
    assert.Reset
    On Error Resume Next
    CollectionErrorTest.TestKeyNotFound assert
    On Error GoTo 0
    If Not assert.Failed Then
        testResults = testResults & "[PASS] TestKeyNotFound" & vbCrLf
        passCount = passCount + 1
    Else
        testResults = testResults & "[FAIL] TestKeyNotFound - " & assert.FailMessage & vbCrLf
        failCount = failCount + 1
        allPass = False
    End If

    ' Execute TestCollectionNoError
    assert.Reset
    On Error Resume Next
    CollectionErrorTest.TestCollectionNoError assert
    On Error GoTo 0
    If Not assert.Failed Then
        testResults = testResults & "[PASS] TestCollectionNoError" & vbCrLf
        passCount = passCount + 1
    Else
        testResults = testResults & "[FAIL] TestCollectionNoError - " & assert.FailMessage & vbCrLf
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
