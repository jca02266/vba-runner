Option Explicit

' VBA の truthy/falsy 仕様確認テスト
'
' VBA の重要な仕様:
'   True  = -1  (1 ではない)
'   False =  0
'   If 条件式は数値として評価: 0 = False, 非ゼロ = True
'   Null を Boolean コンテキストで使うと Error 94
'   Not は算術ビット反転: Not 0 = -1, Not 1 = -2

' ---- True/False の値 ----

Sub Test_TrueValueIsMinusOne(assert)
    assert.Assert CInt(True), -1, "True は -1"
End Sub

Sub Test_FalseValueIsZero(assert)
    assert.Assert CInt(False), 0, "False は 0"
End Sub

Sub Test_TrueIsNotOne(assert)
    Dim result As Boolean
    result = (True = 1)
    assert.IsFalse result, "True = 1 は False (True は -1 であるため)"
End Sub

Sub Test_TrueEqualsMinusOne(assert)
    Dim result As Boolean
    result = (True = -1)
    assert.IsTrue result, "True = -1 は True"
End Sub

' ---- If 条件での truthy/falsy ----

Sub Test_ZeroIsFalsy(assert)
    Dim entered As Boolean
    entered = False
    If 0 Then entered = True
    assert.IsFalse entered, "0 は falsy"
End Sub

Sub Test_NonZeroIsTruthy(assert)
    Dim entered As Boolean
    entered = False
    If 1 Then entered = True
    assert.IsTrue entered, "1 は truthy"
End Sub

Sub Test_MinusOneIsTruthy(assert)
    Dim entered As Boolean
    entered = False
    If -1 Then entered = True
    assert.IsTrue entered, "-1 は truthy"
End Sub

Sub Test_TwoIsTruthy(assert)
    Dim entered As Boolean
    entered = False
    If 2 Then entered = True
    assert.IsTrue entered, "2 は truthy"
End Sub

' ---- 未初期化変数のデフォルト値 ----

Sub Test_UninitBooleanIsFalse(assert)
    Dim b As Boolean
    assert.IsFalse b, "未初期化 Boolean は False"
End Sub

Sub Test_UninitIntegerIsZero(assert)
    Dim n As Integer
    assert.Assert n, 0, "未初期化 Integer は 0"
End Sub

Sub Test_UninitVariantIsEmpty(assert)
    Dim v As Variant
    assert.IsTrue IsEmpty(v), "未初期化 Variant は Empty"
End Sub

' ---- Empty の Boolean コンテキスト ----

Sub Test_EmptyIsFalsy(assert)
    Dim v As Variant
    Dim entered As Boolean
    entered = False
    ' Empty は 0 として扱われる
    If v Then entered = True
    assert.IsFalse entered, "Empty(未初期化Variant) は falsy"
End Sub

Sub Test_EmptyCBoolIsFalse(assert)
    Dim v As Variant
    assert.IsFalse CBool(v), "CBool(Empty) = False"
End Sub

' ---- CBool の変換規則 ----

Sub Test_CBoolZeroIsFalse(assert)
    assert.IsFalse CBool(0), "CBool(0) = False"
End Sub

Sub Test_CBoolOneIsTrue(assert)
    assert.IsTrue CBool(1), "CBool(1) = True"
End Sub

Sub Test_CBoolMinusOneIsTrue(assert)
    assert.IsTrue CBool(-1), "CBool(-1) = True"
End Sub

Sub Test_CBoolTwoIsTrue(assert)
    assert.IsTrue CBool(2), "CBool(2) = True"
End Sub

Sub Test_CBoolStringTrueIsTrue(assert)
    assert.IsTrue CBool("True"), "CBool(""True"") = True"
End Sub

Sub Test_CBoolStringFalseIsFalse(assert)
    assert.IsFalse CBool("False"), "CBool(""False"") = False"
End Sub

Sub Test_CBoolNumericStringZeroIsFalse(assert)
    assert.IsFalse CBool("0"), "CBool(""0"") = False"
End Sub

Sub Test_CBoolNumericStringOneIsTrue(assert)
    assert.IsTrue CBool("1"), "CBool(""1"") = True"
End Sub

' ---- Not 演算子（ビット反転） ----

Sub Test_NotFalseIsTrue(assert)
    assert.IsTrue (Not False), "Not False = True"
End Sub

Sub Test_NotTrueIsFalse(assert)
    assert.IsFalse (Not True), "Not True = False"
End Sub

Sub Test_NotZeroIsMinusOne(assert)
    assert.Assert (Not 0), -1, "Not 0 = -1 (ビット反転)"
End Sub

Sub Test_NotOneIsMinusTwo(assert)
    assert.Assert (Not 1), -2, "Not 1 = -2 (ビット反転)"
End Sub

' ---- Boolean の算術演算 ----

Sub Test_TruePlusTrueIsMinusTwo(assert)
    assert.Assert (True + True), -2, "True + True = -2 (-1 + -1)"
End Sub

Sub Test_TruePlusOneIsZero(assert)
    assert.Assert (True + 1), 0, "True + 1 = 0 (-1 + 1)"
End Sub

Sub Test_FalsePlusFiveIsFive(assert)
    assert.Assert (False + 5), 5, "False + 5 = 5 (0 + 5)"
End Sub

' ---- Null の Boolean コンテキスト (Error 94) ----

Sub Test_NullInIfRaisesError94(assert)
    Dim savedErr As Long
    On Error Resume Next
    Err.Clear
    Dim n As Variant
    n = Null
    Dim entered As Boolean
    entered = False
    If n Then entered = True
    savedErr = Err.Number
    On Error GoTo 0
    assert.Assert savedErr, 94, "Null を If 条件に使うと Error 94"
End Sub

Sub Test_CBoolNullRaisesError94(assert)
    Dim savedErr As Long
    On Error Resume Next
    Err.Clear
    Dim result As Boolean
    result = CBool(Null)
    savedErr = Err.Number
    On Error GoTo 0
    assert.Assert savedErr, 94, "CBool(Null) は Error 94"
End Sub

' ---- If 条件での文字列評価（§5.6.9 + §6.1.2.3.1.1 value coercion） ----

Sub Test_IfStringTrueIsTruthy(assert)
    Dim entered As Boolean
    entered = False
    If "True" Then entered = True
    assert.IsTrue entered, """True"" は truthy"
End Sub

Sub Test_IfStringFalseIsFalsy(assert)
    Dim entered As Boolean
    entered = False
    If "False" Then entered = True
    assert.IsFalse entered, """False"" は falsy"
End Sub

Sub Test_IfStringZeroIsFalsy(assert)
    Dim entered As Boolean
    entered = False
    If "0" Then entered = True
    assert.IsFalse entered, """0"" は falsy（数値 0 に変換）"
End Sub

Sub Test_IfStringOneIsTruthy(assert)
    Dim entered As Boolean
    entered = False
    If "1" Then entered = True
    assert.IsTrue entered, """1"" は truthy（数値 1 に変換）"
End Sub

Sub Test_IfEmptyStringRaisesError13(assert)
    Dim savedErr As Long
    On Error Resume Next
    Err.Clear
    Dim entered As Boolean
    entered = False
    If "" Then entered = True
    savedErr = Err.Number
    On Error GoTo 0
    assert.Assert savedErr, 13, """""""（空文字）を If 条件に使うと Type mismatch Error 13"
End Sub

Sub Test_IfNonNumericStringRaisesError13(assert)
    Dim savedErr As Long
    On Error Resume Next
    Err.Clear
    Dim entered As Boolean
    entered = False
    If "abc" Then entered = True
    savedErr = Err.Number
    On Error GoTo 0
    assert.Assert savedErr, 13, """abc"" を If 条件に使うと Type mismatch Error 13"
End Sub

' ---- CStr による Boolean → 文字列変換 ----

Sub Test_CStrTrueIsStringTrue(assert)
    assert.Assert CStr(True), "True", "CStr(True) = ""True"""
End Sub

Sub Test_CStrFalseIsStringFalse(assert)
    assert.Assert CStr(False), "False", "CStr(False) = ""False"""
End Sub

' ---- 比較演算子の結果は -1 / 0 ----

Sub Test_ComparisonTrueResultIsMinusOne(assert)
    assert.Assert CInt(1 > 0), -1, "CInt(1 > 0) = -1（True = -1）"
End Sub

Sub Test_ComparisonFalseResultIsZero(assert)
    assert.Assert CInt(1 < 0), 0, "CInt(1 < 0) = 0（False = 0）"
End Sub
