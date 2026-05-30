Option Explicit

' Variant 型推論のデモ
' Dim 宣言の右に ⟨Long⟩ などのヒントが表示されるはずです

Function GetPrice() As Long
    GetPrice = 1000
End Function

Function GetLabel()        ' 戻り型未宣言 → ボディから推論
    GetLabel = "商品名"
End Function

Sub VariantDemo()
    ' リテラル代入 → 型が推論される
    Dim count
    count = 42             ' ⟨Long⟩

    Dim price
    price = 3.14           ' ⟨Double⟩

    Dim name
    name = "Alice"         ' ⟨String⟩

    Dim flag
    flag = True            ' ⟨Boolean⟩

    ' 算術演算
    Dim total
    total = count * 2      ' ⟨Long⟩

    ' 文字列連結
    Dim greeting
    greeting = "Hello " & name  ' ⟨String⟩

    ' 明示的な戻り型のある関数呼び出し
    Dim p
    p = GetPrice()         ' ⟨Long⟩

    ' 戻り型未宣言の関数 → ボディを再帰推論
    Dim lbl
    lbl = GetLabel()       ' ⟨String⟩

    ' 組み込み関数
    Dim n
    n = CLng("123")        ' ⟨Long⟩

    Dim s
    s = CStr(42)           ' ⟨String⟩

    Dim b
    b = CBool(1)           ' ⟨Boolean⟩
End Sub
