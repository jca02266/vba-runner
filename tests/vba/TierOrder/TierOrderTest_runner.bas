Option Explicit

' Auto-generated test runner from vba-test-generator
' Run this Sub in Excel VBA environment to execute all tests
' Module-qualified calls: TierOrderTest.TestProcedures
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

    ' Execute Test_Tier4BuiltinAccessible
    assert.Reset
    On Error Resume Next
    TierOrderTest.Test_Tier4BuiltinAccessible assert
    On Error GoTo 0
    If Not assert.Failed Then
        testResults = testResults & "[PASS] Test_Tier4BuiltinAccessible" & vbCrLf
        passCount = passCount + 1
    Else
        testResults = testResults & "[FAIL] Test_Tier4BuiltinAccessible - " & assert.FailMessage & vbCrLf
        failCount = failCount + 1
        allPass = False
    End If

    ' Execute Test_Tier2OverTier4
    assert.Reset
    On Error Resume Next
    TierOrderTest.Test_Tier2OverTier4 assert
    On Error GoTo 0
    If Not assert.Failed Then
        testResults = testResults & "[PASS] Test_Tier2OverTier4" & vbCrLf
        passCount = passCount + 1
    Else
        testResults = testResults & "[FAIL] Test_Tier2OverTier4 - " & assert.FailMessage & vbCrLf
        failCount = failCount + 1
        allPass = False
    End If

    ' Execute Test_Tier3CrossModuleAccess
    assert.Reset
    On Error Resume Next
    TierOrderTest.Test_Tier3CrossModuleAccess assert
    On Error GoTo 0
    If Not assert.Failed Then
        testResults = testResults & "[PASS] Test_Tier3CrossModuleAccess" & vbCrLf
        passCount = passCount + 1
    Else
        testResults = testResults & "[FAIL] Test_Tier3CrossModuleAccess - " & assert.FailMessage & vbCrLf
        failCount = failCount + 1
        allPass = False
    End If

    ' Execute Test_Tier2OverTier3
    assert.Reset
    On Error Resume Next
    TierOrderTest.Test_Tier2OverTier3 assert
    On Error GoTo 0
    If Not assert.Failed Then
        testResults = testResults & "[PASS] Test_Tier2OverTier3" & vbCrLf
        passCount = passCount + 1
    Else
        testResults = testResults & "[FAIL] Test_Tier2OverTier3 - " & assert.FailMessage & vbCrLf
        failCount = failCount + 1
        allPass = False
    End If

    ' Execute Test_Tier1OverTier2
    assert.Reset
    On Error Resume Next
    TierOrderTest.Test_Tier1OverTier2 assert
    On Error GoTo 0
    If Not assert.Failed Then
        testResults = testResults & "[PASS] Test_Tier1OverTier2" & vbCrLf
        passCount = passCount + 1
    Else
        testResults = testResults & "[FAIL] Test_Tier1OverTier2 - " & assert.FailMessage & vbCrLf
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
