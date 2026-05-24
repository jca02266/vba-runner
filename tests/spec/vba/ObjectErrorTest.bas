' オブジェクトエラー番号テスト (Error 91, 424, 438)

' Error 424: Set ステートメントで右辺が非オブジェクト (Set x = 5)
Sub TestSetNonObject(assert)
    Dim savedErr As Long
    On Error Resume Next
    Err.Clear
    Dim obj As Object
    Set obj = 5
    savedErr = Err.Number
    On Error GoTo 0

    assert.Assert savedErr, 424, "Set x = 5 -> Error 424"
End Sub

' Error 424: Set ステートメントで右辺が文字列
Sub TestSetString(assert)
    Dim savedErr As Long
    On Error Resume Next
    Err.Clear
    Dim obj As Object
    Set obj = "hello"
    savedErr = Err.Number
    On Error GoTo 0

    assert.Assert savedErr, 424, "Set x = ""hello"" -> Error 424"
End Sub

' Error 91: Nothing に対してプロパティアクセス
Sub TestAccessNothingProperty(assert)
    Dim savedErr As Long
    On Error Resume Next
    Err.Clear
    Dim obj As Object
    Dim v
    v = obj.SomeProperty
    savedErr = Err.Number
    On Error GoTo 0

    assert.Assert savedErr, 91, "Nothing.Property -> Error 91"
End Sub

' Error 91: Nothing に対してメソッド呼び出し
Sub TestCallNothingMethod(assert)
    Dim savedErr As Long
    On Error Resume Next
    Err.Clear
    Dim obj As Object
    obj.DoSomething
    savedErr = Err.Number
    On Error GoTo 0

    assert.Assert savedErr, 91, "Nothing.Method -> Error 91"
End Sub

' Error 438: クラスに存在しないメソッド呼び出し
Sub TestCallNonExistentMethod(assert)
    Dim savedErr As Long
    On Error Resume Next
    Err.Clear
    Dim col As New Collection
    col.NonExistentMethod
    savedErr = Err.Number
    On Error GoTo 0

    assert.Assert savedErr, 438, "obj.存在しないメソッド -> Error 438"
End Sub

' Error 438: クラスに存在しないプロパティ読み取り
Sub TestReadNonExistentProperty(assert)
    Dim savedErr As Long
    On Error Resume Next
    Err.Clear
    Dim col As New Collection
    Dim v
    v = col.NoSuchProp
    savedErr = Err.Number
    On Error GoTo 0

    assert.Assert savedErr, 438, "obj.存在しないプロパティ -> Error 438"
End Sub

' Error 438: CreateObject 返り値の存在しないメソッド呼び出し (Late Binding)
Sub TestLateBindingNonExistentMethod(assert)
    Dim savedErr As Long
    On Error Resume Next
    Err.Clear
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.NoSuchMethod
    savedErr = Err.Number
    On Error GoTo 0

    assert.Assert savedErr, 438, "CreateObject返り値.存在しないメソッド -> Error 438"
End Sub

' Error 438: CreateObject 返り値の存在しないプロパティ読み取り (Late Binding)
Sub TestLateBindingNonExistentProperty(assert)
    Dim savedErr As Long
    On Error Resume Next
    Err.Clear
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    Dim v
    v = d.NoSuchProperty
    savedErr = Err.Number
    On Error GoTo 0

    assert.Assert savedErr, 438, "CreateObject返り値.存在しないプロパティ -> Error 438"
End Sub
