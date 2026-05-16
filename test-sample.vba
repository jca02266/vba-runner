Option Explicit

' VBA Compiler Extension Test Sample

Public GlobalCounter As Integer

' メイン関数
Function Add(a As Integer, b As Integer) As Integer
    Add = a + b
End Function

' テスト関数
Sub Test_Addition()
    Dim result As Integer
    result = Add(5, 3)
    ' ホバー機能を試す：マウスをresultに乗せるとInteger型が表示される
    ' 定義ジャンプ機能を試す：Addの上でF12を押すと定義に移動
End Sub

' クラスの例
Class Calculator
    Private mTotal As Double

    Sub Add(value As Double)
        mTotal = mTotal + value
    End Sub

    Function GetTotal() As Double
        GetTotal = mTotal
    End Function
End Class

' プロパティの例
Class Person
    Private mName As String
    Private mAge As Integer

    Property Get Name() As String
        Name = mName
    End Property

    Property Let Name(value As String)
        mName = value
    End Property

    Property Get Age() As Integer
        Age = mAge
    End Property

    Property Let Age(value As Integer)
        If value >= 0 Then mAge = value
    End Property
End Class

' テストプロシージャ（VSCode Testing API で実行可能）
Sub Test_Calculator()
    Dim calc As New Calculator
    calc.Add 10
    calc.Add 5
    ' 実行結果: 15
End Sub

Sub Test_Person()
    Dim person As New Person
    person.Name = "Alice"
    person.Age = 30
    ' 実行結果: Person作成完了
End Sub
