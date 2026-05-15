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
    On Error Resume Next
    Err.Clear
    CurrencyOperationsTest.Test_CurrencyArithmetic assert
    If Err.Number = 0 Then
        testResults = testResults & "[PASS] Test_CurrencyArithmetic" & vbCrLf
        passCount = passCount + 1
    Else
        testResults = testResults & "[FAIL] Test_CurrencyArithmetic" & vbCrLf
        failCount = failCount + 1
        allPass = False
    End If
    On Error GoTo 0

    ' Execute Test_CurrencyMultiplication
    On Error Resume Next
    Err.Clear
    CurrencyOperationsTest.Test_CurrencyMultiplication assert
    If Err.Number = 0 Then
        testResults = testResults & "[PASS] Test_CurrencyMultiplication" & vbCrLf
        passCount = passCount + 1
    Else
        testResults = testResults & "[FAIL] Test_CurrencyMultiplication" & vbCrLf
        failCount = failCount + 1
        allPass = False
    End If
    On Error GoTo 0

    ' Execute Test_CurrencyDivision
    On Error Resume Next
    Err.Clear
    CurrencyOperationsTest.Test_CurrencyDivision assert
    If Err.Number = 0 Then
        testResults = testResults & "[PASS] Test_CurrencyDivision" & vbCrLf
        passCount = passCount + 1
    Else
        testResults = testResults & "[FAIL] Test_CurrencyDivision" & vbCrLf
        failCount = failCount + 1
        allPass = False
    End If
    On Error GoTo 0

    ' Execute Test_CurrencyTaxCalculation
    On Error Resume Next
    Err.Clear
    CurrencyOperationsTest.Test_CurrencyTaxCalculation assert
    If Err.Number = 0 Then
        testResults = testResults & "[PASS] Test_CurrencyTaxCalculation" & vbCrLf
        passCount = passCount + 1
    Else
        testResults = testResults & "[FAIL] Test_CurrencyTaxCalculation" & vbCrLf
        failCount = failCount + 1
        allPass = False
    End If
    On Error GoTo 0

    ' Execute Test_CurrencyArray
    On Error Resume Next
    Err.Clear
    CurrencyOperationsTest.Test_CurrencyArray assert
    If Err.Number = 0 Then
        testResults = testResults & "[PASS] Test_CurrencyArray" & vbCrLf
        passCount = passCount + 1
    Else
        testResults = testResults & "[FAIL] Test_CurrencyArray" & vbCrLf
        failCount = failCount + 1
        allPass = False
    End If
    On Error GoTo 0

    ' Execute Test_CurrencyConversion
    On Error Resume Next
    Err.Clear
    CurrencyOperationsTest.Test_CurrencyConversion assert
    If Err.Number = 0 Then
        testResults = testResults & "[PASS] Test_CurrencyConversion" & vbCrLf
        passCount = passCount + 1
    Else
        testResults = testResults & "[FAIL] Test_CurrencyConversion" & vbCrLf
        failCount = failCount + 1
        allPass = False
    End If
    On Error GoTo 0

    ' Execute Test_CurrencyComparison
    On Error Resume Next
    Err.Clear
    CurrencyOperationsTest.Test_CurrencyComparison assert
    If Err.Number = 0 Then
        testResults = testResults & "[PASS] Test_CurrencyComparison" & vbCrLf
        passCount = passCount + 1
    Else
        testResults = testResults & "[FAIL] Test_CurrencyComparison" & vbCrLf
        failCount = failCount + 1
        allPass = False
    End If
    On Error GoTo 0

    ' Execute Test_CurrencyNegative
    On Error Resume Next
    Err.Clear
    CurrencyOperationsTest.Test_CurrencyNegative assert
    If Err.Number = 0 Then
        testResults = testResults & "[PASS] Test_CurrencyNegative" & vbCrLf
        passCount = passCount + 1
    Else
        testResults = testResults & "[FAIL] Test_CurrencyNegative" & vbCrLf
        failCount = failCount + 1
        allPass = False
    End If
    On Error GoTo 0

    ' Execute Test_CurrencyPrecision
    On Error Resume Next
    Err.Clear
    CurrencyOperationsTest.Test_CurrencyPrecision assert
    If Err.Number = 0 Then
        testResults = testResults & "[PASS] Test_CurrencyPrecision" & vbCrLf
        passCount = passCount + 1
    Else
        testResults = testResults & "[FAIL] Test_CurrencyPrecision" & vbCrLf
        failCount = failCount + 1
        allPass = False
    End If
    On Error GoTo 0

    ' Execute Test_CurrencyFunctionParameter
    On Error Resume Next
    Err.Clear
    CurrencyOperationsTest.Test_CurrencyFunctionParameter assert
    If Err.Number = 0 Then
        testResults = testResults & "[PASS] Test_CurrencyFunctionParameter" & vbCrLf
        passCount = passCount + 1
    Else
        testResults = testResults & "[FAIL] Test_CurrencyFunctionParameter" & vbCrLf
        failCount = failCount + 1
        allPass = False
    End If
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
