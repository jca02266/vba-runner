' Inlay Hint 表示位置確認用サンプル
' VSCode で開いて Inlay Hint が () の後に表示されることを確認する

' 1. 戻り型未宣言の関数 → ボディから推論
'    修正前の表示: Function GetLabel As String()
'    修正後の表示: Function GetLabel() As String
Function GetLabel()
    GetLabel = "商品名"
End Function

Function GetCount()
    GetCount = 42
End Function

' 2. パラメーター・変数・動的配列の型推論
'    修正前の表示: arr As String()
'    修正後の表示: arr() As String
Sub InlayHintPositionDemo(name, count)
    ' ローカル変数推論
    Dim label
    label = GetLabel()      ' label As String

    Dim total
    total = GetCount() * 2  ' total As Long

    ' 動的配列宣言 → ヒントは arr() の後に表示
    Dim arr()
    arr = UCase("a,b,c")    ' arr() As String
End Sub
