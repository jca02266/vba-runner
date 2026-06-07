Option Explicit

' Auto-generated test runner from vba-test-generator
' Run this Sub in Excel VBA environment to execute all tests
' Module-qualified calls: CurrencyOperationsTest.TestProcedures
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

    ' Execute Test_CurrencyArithmetic
    assert.Reset
    On Error Resume Next
    CurrencyOperationsTest.Test_CurrencyArithmetic assert
    On Error GoTo 0
    If Not assert.Failed Then
        testResults = testResults & "[PASS] Test_CurrencyArithmetic" & vbCrLf
        passCount = passCount + 1
    Else
        testResults = testResults & "[FAIL] Test_CurrencyArithmetic - " & assert.FailMessage & vbCrLf
        failCount = failCount + 1
        allPass = False
    End If

    ' Execute Test_CurrencyMultiplication
    assert.Reset
    On Error Resume Next
    CurrencyOperationsTest.Test_CurrencyMultiplication assert
    On Error GoTo 0
    If Not assert.Failed Then
        testResults = testResults & "[PASS] Test_CurrencyMultiplication" & vbCrLf
        passCount = passCount + 1
    Else
        testResults = testResults & "[FAIL] Test_CurrencyMultiplication - " & assert.FailMessage & vbCrLf
        failCount = failCount + 1
        allPass = False
    End If

    ' Execute Test_CurrencyDivision
    assert.Reset
    On Error Resume Next
    CurrencyOperationsTest.Test_CurrencyDivision assert
    On Error GoTo 0
    If Not assert.Failed Then
        testResults = testResults & "[PASS] Test_CurrencyDivision" & vbCrLf
        passCount = passCount + 1
    Else
        testResults = testResults & "[FAIL] Test_CurrencyDivision - " & assert.FailMessage & vbCrLf
        failCount = failCount + 1
        allPass = False
    End If

    ' Execute Test_CurrencyTaxCalculation
    assert.Reset
    On Error Resume Next
    CurrencyOperationsTest.Test_CurrencyTaxCalculation assert
    On Error GoTo 0
    If Not assert.Failed Then
        testResults = testResults & "[PASS] Test_CurrencyTaxCalculation" & vbCrLf
        passCount = passCount + 1
    Else
        testResults = testResults & "[FAIL] Test_CurrencyTaxCalculation - " & assert.FailMessage & vbCrLf
        failCount = failCount + 1
        allPass = False
    End If

    ' Execute Test_CurrencyArray
    assert.Reset
    On Error Resume Next
    CurrencyOperationsTest.Test_CurrencyArray assert
    On Error GoTo 0
    If Not assert.Failed Then
        testResults = testResults & "[PASS] Test_CurrencyArray" & vbCrLf
        passCount = passCount + 1
    Else
        testResults = testResults & "[FAIL] Test_CurrencyArray - " & assert.FailMessage & vbCrLf
        failCount = failCount + 1
        allPass = False
    End If

    ' Execute Test_CurrencyConversion
    assert.Reset
    On Error Resume Next
    CurrencyOperationsTest.Test_CurrencyConversion assert
    On Error GoTo 0
    If Not assert.Failed Then
        testResults = testResults & "[PASS] Test_CurrencyConversion" & vbCrLf
        passCount = passCount + 1
    Else
        testResults = testResults & "[FAIL] Test_CurrencyConversion - " & assert.FailMessage & vbCrLf
        failCount = failCount + 1
        allPass = False
    End If

    ' Execute Test_CurrencyComparison
    assert.Reset
    On Error Resume Next
    CurrencyOperationsTest.Test_CurrencyComparison assert
    On Error GoTo 0
    If Not assert.Failed Then
        testResults = testResults & "[PASS] Test_CurrencyComparison" & vbCrLf
        passCount = passCount + 1
    Else
        testResults = testResults & "[FAIL] Test_CurrencyComparison - " & assert.FailMessage & vbCrLf
        failCount = failCount + 1
        allPass = False
    End If

    ' Execute Test_CurrencyNegative
    assert.Reset
    On Error Resume Next
    CurrencyOperationsTest.Test_CurrencyNegative assert
    On Error GoTo 0
    If Not assert.Failed Then
        testResults = testResults & "[PASS] Test_CurrencyNegative" & vbCrLf
        passCount = passCount + 1
    Else
        testResults = testResults & "[FAIL] Test_CurrencyNegative - " & assert.FailMessage & vbCrLf
        failCount = failCount + 1
        allPass = False
    End If

    ' Execute Test_CurrencyPrecision
    assert.Reset
    On Error Resume Next
    CurrencyOperationsTest.Test_CurrencyPrecision assert
    On Error GoTo 0
    If Not assert.Failed Then
        testResults = testResults & "[PASS] Test_CurrencyPrecision" & vbCrLf
        passCount = passCount + 1
    Else
        testResults = testResults & "[FAIL] Test_CurrencyPrecision - " & assert.FailMessage & vbCrLf
        failCount = failCount + 1
        allPass = False
    End If

    ' Execute Test_CurrencyFunctionParameter
    assert.Reset
    On Error Resume Next
    CurrencyOperationsTest.Test_CurrencyFunctionParameter assert
    On Error GoTo 0
    If Not assert.Failed Then
        testResults = testResults & "[PASS] Test_CurrencyFunctionParameter" & vbCrLf
        passCount = passCount + 1
    Else
        testResults = testResults & "[FAIL] Test_CurrencyFunctionParameter - " & assert.FailMessage & vbCrLf
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
