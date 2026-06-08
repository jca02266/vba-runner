Option Explicit

' Auto-generated test runner from vba-test-generator
' Run this Sub in Excel VBA environment to execute all tests
' Module-qualified calls: Tier5Test.TestProcedures
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

    ' Execute Test_LocalVarShadowsComClass
    assert.Reset
    On Error Resume Next
    Tier5Test.Test_LocalVarShadowsComClass assert
    On Error GoTo 0
    If Not assert.Failed Then
        testResults = testResults & "[PASS] Test_LocalVarShadowsComClass" & vbCrLf
        passCount = passCount + 1
    Else
        testResults = testResults & "[FAIL] Test_LocalVarShadowsComClass - " & assert.FailMessage & vbCrLf
        failCount = failCount + 1
        allPass = False
    End If

    ' Execute Test_LocalVarShadowsComNamespace
    assert.Reset
    On Error Resume Next
    Tier5Test.Test_LocalVarShadowsComNamespace assert
    On Error GoTo 0
    If Not assert.Failed Then
        testResults = testResults & "[PASS] Test_LocalVarShadowsComNamespace" & vbCrLf
        passCount = passCount + 1
    Else
        testResults = testResults & "[FAIL] Test_LocalVarShadowsComNamespace - " & assert.FailMessage & vbCrLf
        failCount = failCount + 1
        allPass = False
    End If

    ' Execute Test_ComObjectAccessibleViaCreateObject
    assert.Reset
    On Error Resume Next
    Tier5Test.Test_ComObjectAccessibleViaCreateObject assert
    On Error GoTo 0
    If Not assert.Failed Then
        testResults = testResults & "[PASS] Test_ComObjectAccessibleViaCreateObject" & vbCrLf
        passCount = passCount + 1
    Else
        testResults = testResults & "[FAIL] Test_ComObjectAccessibleViaCreateObject - " & assert.FailMessage & vbCrLf
        failCount = failCount + 1
        allPass = False
    End If

    ' Execute Test_ShortClassNameAccessible
    assert.Reset
    On Error Resume Next
    Tier5Test.Test_ShortClassNameAccessible assert
    On Error GoTo 0
    If Not assert.Failed Then
        testResults = testResults & "[PASS] Test_ShortClassNameAccessible" & vbCrLf
        passCount = passCount + 1
    Else
        testResults = testResults & "[FAIL] Test_ShortClassNameAccessible - " & assert.FailMessage & vbCrLf
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
