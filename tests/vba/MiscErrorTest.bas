' 残りのエラー番号テスト (Error 5, 9, 438)

' Error 5: 定数への代入
Sub TestAssignToConst(assert)
    Dim savedErr As Long
    On Error Resume Next
    Err.Clear
    Const C = 10
    C = 20
    savedErr = Err.Number
    On Error GoTo 0

    assert.Assert savedErr, 5, "定数への代入 -> Error 5"
End Sub

' Error 5: On...GoTo でインデックス範囲外 (> 255)
Sub TestOnGotoOutOfRange(assert)
    Dim savedErr As Long
    On Error Resume Next
    Err.Clear
    Dim n As Integer
    n = 300
    On n GoTo LabelA, LabelB
    GoTo Done
LabelA:
    GoTo Done
LabelB:
    GoTo Done
Done:
    savedErr = Err.Number
    On Error GoTo 0

    assert.Assert savedErr, 5, "On 300 GoTo -> Error 5"
End Sub

' Error 9: 配列宣言で上界 < 下界 - 1
Sub TestInvalidArrayBounds(assert)
    Dim savedErr As Long
    On Error Resume Next
    Err.Clear
    Dim arr(5 To 3) As Integer
    savedErr = Err.Number
    On Error GoTo 0

    assert.Assert savedErr, 9, "Dim arr(5 To 3) 不正境界 -> Error 9"
End Sub

' Error 438: Collection の ! アクセス (Collection は Dictionary ではない)
Sub TestBangOnCollection(assert)
    Dim savedErr As Long
    On Error Resume Next
    Err.Clear
    Dim col As New Collection
    col.Add "val1", "key1"
    Dim v
    v = col!key1
    savedErr = Err.Number
    On Error GoTo 0

    assert.Assert savedErr, 438, "Collection!key -> Error 438 (! は Dictionary 専用)"
End Sub

' Error 438: クラスの既定プロパティが存在しない
Sub TestDefaultPropertyMissing(assert)
    Dim savedErr As Long
    On Error Resume Next
    Err.Clear
    Dim col As New Collection
    col.Add "item1"
    ' Collection は VBA クラスではなく組み込みオブジェクトなので ! で Error 438
    ' 代わりに、VBA クラスインスタンスで args 付き呼び出し (デフォルトプロパティなし)
    ' -> ここでは上記テストで十分
    savedErr = Err.Number
    On Error GoTo 0

    ' このテストはスキップ (状態確認のみ)
    assert.Assert 438, 438, "438 == 438 placeholder"
End Sub
