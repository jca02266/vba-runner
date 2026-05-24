Option Explicit

' Excel VBA 環境でのテスト実行ランナー
' TruthyFalsyTest.bas の全 Test_* Sub を実行する
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

    ' Execute Test_TrueValueIsMinusOne
    assert.Reset
    On Error Resume Next
    TruthyFalsyTest.Test_TrueValueIsMinusOne assert
    On Error GoTo 0
    If Not assert.Failed Then
        testResults = testResults & "[PASS] Test_TrueValueIsMinusOne" & vbCrLf
        passCount = passCount + 1
    Else
        testResults = testResults & "[FAIL] Test_TrueValueIsMinusOne - " & assert.FailMessage & vbCrLf
        failCount = failCount + 1
        allPass = False
    End If

    ' Execute Test_FalseValueIsZero
    assert.Reset
    On Error Resume Next
    TruthyFalsyTest.Test_FalseValueIsZero assert
    On Error GoTo 0
    If Not assert.Failed Then
        testResults = testResults & "[PASS] Test_FalseValueIsZero" & vbCrLf
        passCount = passCount + 1
    Else
        testResults = testResults & "[FAIL] Test_FalseValueIsZero - " & assert.FailMessage & vbCrLf
        failCount = failCount + 1
        allPass = False
    End If

    ' Execute Test_TrueIsNotOne
    assert.Reset
    On Error Resume Next
    TruthyFalsyTest.Test_TrueIsNotOne assert
    On Error GoTo 0
    If Not assert.Failed Then
        testResults = testResults & "[PASS] Test_TrueIsNotOne" & vbCrLf
        passCount = passCount + 1
    Else
        testResults = testResults & "[FAIL] Test_TrueIsNotOne - " & assert.FailMessage & vbCrLf
        failCount = failCount + 1
        allPass = False
    End If

    ' Execute Test_TrueEqualsMinusOne
    assert.Reset
    On Error Resume Next
    TruthyFalsyTest.Test_TrueEqualsMinusOne assert
    On Error GoTo 0
    If Not assert.Failed Then
        testResults = testResults & "[PASS] Test_TrueEqualsMinusOne" & vbCrLf
        passCount = passCount + 1
    Else
        testResults = testResults & "[FAIL] Test_TrueEqualsMinusOne - " & assert.FailMessage & vbCrLf
        failCount = failCount + 1
        allPass = False
    End If

    ' Execute Test_ZeroIsFalsy
    assert.Reset
    On Error Resume Next
    TruthyFalsyTest.Test_ZeroIsFalsy assert
    On Error GoTo 0
    If Not assert.Failed Then
        testResults = testResults & "[PASS] Test_ZeroIsFalsy" & vbCrLf
        passCount = passCount + 1
    Else
        testResults = testResults & "[FAIL] Test_ZeroIsFalsy - " & assert.FailMessage & vbCrLf
        failCount = failCount + 1
        allPass = False
    End If

    ' Execute Test_NonZeroIsTruthy
    assert.Reset
    On Error Resume Next
    TruthyFalsyTest.Test_NonZeroIsTruthy assert
    On Error GoTo 0
    If Not assert.Failed Then
        testResults = testResults & "[PASS] Test_NonZeroIsTruthy" & vbCrLf
        passCount = passCount + 1
    Else
        testResults = testResults & "[FAIL] Test_NonZeroIsTruthy - " & assert.FailMessage & vbCrLf
        failCount = failCount + 1
        allPass = False
    End If

    ' Execute Test_MinusOneIsTruthy
    assert.Reset
    On Error Resume Next
    TruthyFalsyTest.Test_MinusOneIsTruthy assert
    On Error GoTo 0
    If Not assert.Failed Then
        testResults = testResults & "[PASS] Test_MinusOneIsTruthy" & vbCrLf
        passCount = passCount + 1
    Else
        testResults = testResults & "[FAIL] Test_MinusOneIsTruthy - " & assert.FailMessage & vbCrLf
        failCount = failCount + 1
        allPass = False
    End If

    ' Execute Test_TwoIsTruthy
    assert.Reset
    On Error Resume Next
    TruthyFalsyTest.Test_TwoIsTruthy assert
    On Error GoTo 0
    If Not assert.Failed Then
        testResults = testResults & "[PASS] Test_TwoIsTruthy" & vbCrLf
        passCount = passCount + 1
    Else
        testResults = testResults & "[FAIL] Test_TwoIsTruthy - " & assert.FailMessage & vbCrLf
        failCount = failCount + 1
        allPass = False
    End If

    ' Execute Test_UninitBooleanIsFalse
    assert.Reset
    On Error Resume Next
    TruthyFalsyTest.Test_UninitBooleanIsFalse assert
    On Error GoTo 0
    If Not assert.Failed Then
        testResults = testResults & "[PASS] Test_UninitBooleanIsFalse" & vbCrLf
        passCount = passCount + 1
    Else
        testResults = testResults & "[FAIL] Test_UninitBooleanIsFalse - " & assert.FailMessage & vbCrLf
        failCount = failCount + 1
        allPass = False
    End If

    ' Execute Test_UninitIntegerIsZero
    assert.Reset
    On Error Resume Next
    TruthyFalsyTest.Test_UninitIntegerIsZero assert
    On Error GoTo 0
    If Not assert.Failed Then
        testResults = testResults & "[PASS] Test_UninitIntegerIsZero" & vbCrLf
        passCount = passCount + 1
    Else
        testResults = testResults & "[FAIL] Test_UninitIntegerIsZero - " & assert.FailMessage & vbCrLf
        failCount = failCount + 1
        allPass = False
    End If

    ' Execute Test_UninitVariantIsEmpty
    assert.Reset
    On Error Resume Next
    TruthyFalsyTest.Test_UninitVariantIsEmpty assert
    On Error GoTo 0
    If Not assert.Failed Then
        testResults = testResults & "[PASS] Test_UninitVariantIsEmpty" & vbCrLf
        passCount = passCount + 1
    Else
        testResults = testResults & "[FAIL] Test_UninitVariantIsEmpty - " & assert.FailMessage & vbCrLf
        failCount = failCount + 1
        allPass = False
    End If

    ' Execute Test_EmptyIsFalsy
    assert.Reset
    On Error Resume Next
    TruthyFalsyTest.Test_EmptyIsFalsy assert
    On Error GoTo 0
    If Not assert.Failed Then
        testResults = testResults & "[PASS] Test_EmptyIsFalsy" & vbCrLf
        passCount = passCount + 1
    Else
        testResults = testResults & "[FAIL] Test_EmptyIsFalsy - " & assert.FailMessage & vbCrLf
        failCount = failCount + 1
        allPass = False
    End If

    ' Execute Test_EmptyCBoolIsFalse
    assert.Reset
    On Error Resume Next
    TruthyFalsyTest.Test_EmptyCBoolIsFalse assert
    On Error GoTo 0
    If Not assert.Failed Then
        testResults = testResults & "[PASS] Test_EmptyCBoolIsFalse" & vbCrLf
        passCount = passCount + 1
    Else
        testResults = testResults & "[FAIL] Test_EmptyCBoolIsFalse - " & assert.FailMessage & vbCrLf
        failCount = failCount + 1
        allPass = False
    End If

    ' Execute Test_CBoolZeroIsFalse
    assert.Reset
    On Error Resume Next
    TruthyFalsyTest.Test_CBoolZeroIsFalse assert
    On Error GoTo 0
    If Not assert.Failed Then
        testResults = testResults & "[PASS] Test_CBoolZeroIsFalse" & vbCrLf
        passCount = passCount + 1
    Else
        testResults = testResults & "[FAIL] Test_CBoolZeroIsFalse - " & assert.FailMessage & vbCrLf
        failCount = failCount + 1
        allPass = False
    End If

    ' Execute Test_CBoolOneIsTrue
    assert.Reset
    On Error Resume Next
    TruthyFalsyTest.Test_CBoolOneIsTrue assert
    On Error GoTo 0
    If Not assert.Failed Then
        testResults = testResults & "[PASS] Test_CBoolOneIsTrue" & vbCrLf
        passCount = passCount + 1
    Else
        testResults = testResults & "[FAIL] Test_CBoolOneIsTrue - " & assert.FailMessage & vbCrLf
        failCount = failCount + 1
        allPass = False
    End If

    ' Execute Test_CBoolMinusOneIsTrue
    assert.Reset
    On Error Resume Next
    TruthyFalsyTest.Test_CBoolMinusOneIsTrue assert
    On Error GoTo 0
    If Not assert.Failed Then
        testResults = testResults & "[PASS] Test_CBoolMinusOneIsTrue" & vbCrLf
        passCount = passCount + 1
    Else
        testResults = testResults & "[FAIL] Test_CBoolMinusOneIsTrue - " & assert.FailMessage & vbCrLf
        failCount = failCount + 1
        allPass = False
    End If

    ' Execute Test_CBoolTwoIsTrue
    assert.Reset
    On Error Resume Next
    TruthyFalsyTest.Test_CBoolTwoIsTrue assert
    On Error GoTo 0
    If Not assert.Failed Then
        testResults = testResults & "[PASS] Test_CBoolTwoIsTrue" & vbCrLf
        passCount = passCount + 1
    Else
        testResults = testResults & "[FAIL] Test_CBoolTwoIsTrue - " & assert.FailMessage & vbCrLf
        failCount = failCount + 1
        allPass = False
    End If

    ' Execute Test_CBoolStringTrueIsTrue
    assert.Reset
    On Error Resume Next
    TruthyFalsyTest.Test_CBoolStringTrueIsTrue assert
    On Error GoTo 0
    If Not assert.Failed Then
        testResults = testResults & "[PASS] Test_CBoolStringTrueIsTrue" & vbCrLf
        passCount = passCount + 1
    Else
        testResults = testResults & "[FAIL] Test_CBoolStringTrueIsTrue - " & assert.FailMessage & vbCrLf
        failCount = failCount + 1
        allPass = False
    End If

    ' Execute Test_CBoolStringFalseIsFalse
    assert.Reset
    On Error Resume Next
    TruthyFalsyTest.Test_CBoolStringFalseIsFalse assert
    On Error GoTo 0
    If Not assert.Failed Then
        testResults = testResults & "[PASS] Test_CBoolStringFalseIsFalse" & vbCrLf
        passCount = passCount + 1
    Else
        testResults = testResults & "[FAIL] Test_CBoolStringFalseIsFalse - " & assert.FailMessage & vbCrLf
        failCount = failCount + 1
        allPass = False
    End If

    ' Execute Test_CBoolNumericStringZeroIsFalse
    assert.Reset
    On Error Resume Next
    TruthyFalsyTest.Test_CBoolNumericStringZeroIsFalse assert
    On Error GoTo 0
    If Not assert.Failed Then
        testResults = testResults & "[PASS] Test_CBoolNumericStringZeroIsFalse" & vbCrLf
        passCount = passCount + 1
    Else
        testResults = testResults & "[FAIL] Test_CBoolNumericStringZeroIsFalse - " & assert.FailMessage & vbCrLf
        failCount = failCount + 1
        allPass = False
    End If

    ' Execute Test_CBoolNumericStringOneIsTrue
    assert.Reset
    On Error Resume Next
    TruthyFalsyTest.Test_CBoolNumericStringOneIsTrue assert
    On Error GoTo 0
    If Not assert.Failed Then
        testResults = testResults & "[PASS] Test_CBoolNumericStringOneIsTrue" & vbCrLf
        passCount = passCount + 1
    Else
        testResults = testResults & "[FAIL] Test_CBoolNumericStringOneIsTrue - " & assert.FailMessage & vbCrLf
        failCount = failCount + 1
        allPass = False
    End If

    ' Execute Test_NotFalseIsTrue
    assert.Reset
    On Error Resume Next
    TruthyFalsyTest.Test_NotFalseIsTrue assert
    On Error GoTo 0
    If Not assert.Failed Then
        testResults = testResults & "[PASS] Test_NotFalseIsTrue" & vbCrLf
        passCount = passCount + 1
    Else
        testResults = testResults & "[FAIL] Test_NotFalseIsTrue - " & assert.FailMessage & vbCrLf
        failCount = failCount + 1
        allPass = False
    End If

    ' Execute Test_NotTrueIsFalse
    assert.Reset
    On Error Resume Next
    TruthyFalsyTest.Test_NotTrueIsFalse assert
    On Error GoTo 0
    If Not assert.Failed Then
        testResults = testResults & "[PASS] Test_NotTrueIsFalse" & vbCrLf
        passCount = passCount + 1
    Else
        testResults = testResults & "[FAIL] Test_NotTrueIsFalse - " & assert.FailMessage & vbCrLf
        failCount = failCount + 1
        allPass = False
    End If

    ' Execute Test_NotZeroIsMinusOne
    assert.Reset
    On Error Resume Next
    TruthyFalsyTest.Test_NotZeroIsMinusOne assert
    On Error GoTo 0
    If Not assert.Failed Then
        testResults = testResults & "[PASS] Test_NotZeroIsMinusOne" & vbCrLf
        passCount = passCount + 1
    Else
        testResults = testResults & "[FAIL] Test_NotZeroIsMinusOne - " & assert.FailMessage & vbCrLf
        failCount = failCount + 1
        allPass = False
    End If

    ' Execute Test_NotOneIsMinusTwo
    assert.Reset
    On Error Resume Next
    TruthyFalsyTest.Test_NotOneIsMinusTwo assert
    On Error GoTo 0
    If Not assert.Failed Then
        testResults = testResults & "[PASS] Test_NotOneIsMinusTwo" & vbCrLf
        passCount = passCount + 1
    Else
        testResults = testResults & "[FAIL] Test_NotOneIsMinusTwo - " & assert.FailMessage & vbCrLf
        failCount = failCount + 1
        allPass = False
    End If

    ' Execute Test_TruePlusTrueIsMinusTwo
    assert.Reset
    On Error Resume Next
    TruthyFalsyTest.Test_TruePlusTrueIsMinusTwo assert
    On Error GoTo 0
    If Not assert.Failed Then
        testResults = testResults & "[PASS] Test_TruePlusTrueIsMinusTwo" & vbCrLf
        passCount = passCount + 1
    Else
        testResults = testResults & "[FAIL] Test_TruePlusTrueIsMinusTwo - " & assert.FailMessage & vbCrLf
        failCount = failCount + 1
        allPass = False
    End If

    ' Execute Test_TruePlusOneIsZero
    assert.Reset
    On Error Resume Next
    TruthyFalsyTest.Test_TruePlusOneIsZero assert
    On Error GoTo 0
    If Not assert.Failed Then
        testResults = testResults & "[PASS] Test_TruePlusOneIsZero" & vbCrLf
        passCount = passCount + 1
    Else
        testResults = testResults & "[FAIL] Test_TruePlusOneIsZero - " & assert.FailMessage & vbCrLf
        failCount = failCount + 1
        allPass = False
    End If

    ' Execute Test_FalsePlusFiveIsFive
    assert.Reset
    On Error Resume Next
    TruthyFalsyTest.Test_FalsePlusFiveIsFive assert
    On Error GoTo 0
    If Not assert.Failed Then
        testResults = testResults & "[PASS] Test_FalsePlusFiveIsFive" & vbCrLf
        passCount = passCount + 1
    Else
        testResults = testResults & "[FAIL] Test_FalsePlusFiveIsFive - " & assert.FailMessage & vbCrLf
        failCount = failCount + 1
        allPass = False
    End If

    ' Execute Test_NullInIfRaisesError94
    assert.Reset
    On Error Resume Next
    TruthyFalsyTest.Test_NullInIfRaisesError94 assert
    On Error GoTo 0
    If Not assert.Failed Then
        testResults = testResults & "[PASS] Test_NullInIfRaisesError94" & vbCrLf
        passCount = passCount + 1
    Else
        testResults = testResults & "[FAIL] Test_NullInIfRaisesError94 - " & assert.FailMessage & vbCrLf
        failCount = failCount + 1
        allPass = False
    End If

    ' Execute Test_CBoolNullRaisesError94
    assert.Reset
    On Error Resume Next
    TruthyFalsyTest.Test_CBoolNullRaisesError94 assert
    On Error GoTo 0
    If Not assert.Failed Then
        testResults = testResults & "[PASS] Test_CBoolNullRaisesError94" & vbCrLf
        passCount = passCount + 1
    Else
        testResults = testResults & "[FAIL] Test_CBoolNullRaisesError94 - " & assert.FailMessage & vbCrLf
        failCount = failCount + 1
        allPass = False
    End If

    ' Execute Test_IfStringTrueIsTruthy
    assert.Reset
    On Error Resume Next
    TruthyFalsyTest.Test_IfStringTrueIsTruthy assert
    On Error GoTo 0
    If Not assert.Failed Then
        testResults = testResults & "[PASS] Test_IfStringTrueIsTruthy" & vbCrLf
        passCount = passCount + 1
    Else
        testResults = testResults & "[FAIL] Test_IfStringTrueIsTruthy - " & assert.FailMessage & vbCrLf
        failCount = failCount + 1
        allPass = False
    End If

    ' Execute Test_IfStringFalseIsFalsy
    assert.Reset
    On Error Resume Next
    TruthyFalsyTest.Test_IfStringFalseIsFalsy assert
    On Error GoTo 0
    If Not assert.Failed Then
        testResults = testResults & "[PASS] Test_IfStringFalseIsFalsy" & vbCrLf
        passCount = passCount + 1
    Else
        testResults = testResults & "[FAIL] Test_IfStringFalseIsFalsy - " & assert.FailMessage & vbCrLf
        failCount = failCount + 1
        allPass = False
    End If

    ' Execute Test_IfStringZeroIsFalsy
    assert.Reset
    On Error Resume Next
    TruthyFalsyTest.Test_IfStringZeroIsFalsy assert
    On Error GoTo 0
    If Not assert.Failed Then
        testResults = testResults & "[PASS] Test_IfStringZeroIsFalsy" & vbCrLf
        passCount = passCount + 1
    Else
        testResults = testResults & "[FAIL] Test_IfStringZeroIsFalsy - " & assert.FailMessage & vbCrLf
        failCount = failCount + 1
        allPass = False
    End If

    ' Execute Test_IfStringOneIsTruthy
    assert.Reset
    On Error Resume Next
    TruthyFalsyTest.Test_IfStringOneIsTruthy assert
    On Error GoTo 0
    If Not assert.Failed Then
        testResults = testResults & "[PASS] Test_IfStringOneIsTruthy" & vbCrLf
        passCount = passCount + 1
    Else
        testResults = testResults & "[FAIL] Test_IfStringOneIsTruthy - " & assert.FailMessage & vbCrLf
        failCount = failCount + 1
        allPass = False
    End If

    ' Execute Test_IfEmptyStringRaisesError13
    assert.Reset
    On Error Resume Next
    TruthyFalsyTest.Test_IfEmptyStringRaisesError13 assert
    On Error GoTo 0
    If Not assert.Failed Then
        testResults = testResults & "[PASS] Test_IfEmptyStringRaisesError13" & vbCrLf
        passCount = passCount + 1
    Else
        testResults = testResults & "[FAIL] Test_IfEmptyStringRaisesError13 - " & assert.FailMessage & vbCrLf
        failCount = failCount + 1
        allPass = False
    End If

    ' Execute Test_IfNonNumericStringRaisesError13
    assert.Reset
    On Error Resume Next
    TruthyFalsyTest.Test_IfNonNumericStringRaisesError13 assert
    On Error GoTo 0
    If Not assert.Failed Then
        testResults = testResults & "[PASS] Test_IfNonNumericStringRaisesError13" & vbCrLf
        passCount = passCount + 1
    Else
        testResults = testResults & "[FAIL] Test_IfNonNumericStringRaisesError13 - " & assert.FailMessage & vbCrLf
        failCount = failCount + 1
        allPass = False
    End If

    ' Execute Test_CStrTrueIsStringTrue
    assert.Reset
    On Error Resume Next
    TruthyFalsyTest.Test_CStrTrueIsStringTrue assert
    On Error GoTo 0
    If Not assert.Failed Then
        testResults = testResults & "[PASS] Test_CStrTrueIsStringTrue" & vbCrLf
        passCount = passCount + 1
    Else
        testResults = testResults & "[FAIL] Test_CStrTrueIsStringTrue - " & assert.FailMessage & vbCrLf
        failCount = failCount + 1
        allPass = False
    End If

    ' Execute Test_CStrFalseIsStringFalse
    assert.Reset
    On Error Resume Next
    TruthyFalsyTest.Test_CStrFalseIsStringFalse assert
    On Error GoTo 0
    If Not assert.Failed Then
        testResults = testResults & "[PASS] Test_CStrFalseIsStringFalse" & vbCrLf
        passCount = passCount + 1
    Else
        testResults = testResults & "[FAIL] Test_CStrFalseIsStringFalse - " & assert.FailMessage & vbCrLf
        failCount = failCount + 1
        allPass = False
    End If

    ' Execute Test_ComparisonTrueResultIsMinusOne
    assert.Reset
    On Error Resume Next
    TruthyFalsyTest.Test_ComparisonTrueResultIsMinusOne assert
    On Error GoTo 0
    If Not assert.Failed Then
        testResults = testResults & "[PASS] Test_ComparisonTrueResultIsMinusOne" & vbCrLf
        passCount = passCount + 1
    Else
        testResults = testResults & "[FAIL] Test_ComparisonTrueResultIsMinusOne - " & assert.FailMessage & vbCrLf
        failCount = failCount + 1
        allPass = False
    End If

    ' Execute Test_ComparisonFalseResultIsZero
    assert.Reset
    On Error Resume Next
    TruthyFalsyTest.Test_ComparisonFalseResultIsZero assert
    On Error GoTo 0
    If Not assert.Failed Then
        testResults = testResults & "[PASS] Test_ComparisonFalseResultIsZero" & vbCrLf
        passCount = passCount + 1
    Else
        testResults = testResults & "[FAIL] Test_ComparisonFalseResultIsZero - " & assert.FailMessage & vbCrLf
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
