' Collection のエラー番号テスト (Error 5, 9, 457)

' Error 457: 重複キーで Add
Sub TestDuplicateKeyError(assert)
    Dim col As New Collection
    col.Add "first", "key1"

    Dim savedErr As Long
    On Error Resume Next
    Err.Clear
    col.Add "second", "key1"
    savedErr = Err.Number
    On Error GoTo 0

    assert.Assert savedErr, 457, "重複キー Add -> Error 457"
End Sub

' Error 9: 数値インデックスが範囲外
Sub TestNumericIndexOutOfRange(assert)
    Dim col As New Collection
    col.Add "item1"

    Dim savedErr As Long
    On Error Resume Next
    Err.Clear
    Dim v
    v = col(5)
    savedErr = Err.Number
    On Error GoTo 0

    assert.Assert savedErr, 9, "範囲外インデックス -> Error 9"
End Sub

' Error 9: Remove で範囲外インデックス
Sub TestRemoveOutOfRange(assert)
    Dim col As New Collection
    col.Add "item1"

    Dim savedErr As Long
    On Error Resume Next
    Err.Clear
    col.Remove 99
    savedErr = Err.Number
    On Error GoTo 0

    assert.Assert savedErr, 9, "Remove 範囲外 -> Error 9"
End Sub

' Error 5: 存在しないキー文字列でアクセス
Sub TestKeyNotFound(assert)
    Dim col As New Collection
    col.Add "item1", "key1"

    Dim savedErr As Long
    On Error Resume Next
    Err.Clear
    Dim v
    v = col("nonexistent")
    savedErr = Err.Number
    On Error GoTo 0

    assert.Assert savedErr, 5, "存在しないキー -> Error 5"
End Sub

' 正常系: エラーなし
Sub TestCollectionNoError(assert)
    Dim col As New Collection
    col.Add "a", "k1"
    col.Add "b", "k2"
    col.Add "c", "k3"

    assert.Assert col.Count, 3, "Count = 3"
    assert.Assert col("k2"), "b", "Item by key"
    assert.Assert col(1), "a", "Item by index 1"
    assert.Assert col(3), "c", "Item by index 3"
End Sub
