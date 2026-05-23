' ============================================
' リファクタリング後：テーブル駆動パターン
' 承認ルールをデータテーブルで定義
' ============================================

' 承認ルール定義型
Type ApprovalRule
    department As String
    threshold1 As Long
    approver1 As String
    threshold2 As Long
    approver2 As String
    threshold3 As Long
    approver3 As String
    defaultApprover As String
End Type

' グローバル ルールテーブル
Dim g_rules() As ApprovalRule

' 初期化：ルールテーブルを構築
Sub InitializeApprovalRules()
    ReDim g_rules(4)

    ' Sales部
    g_rules(0).department = "Sales"
    g_rules(0).threshold1 = 50000
    g_rules(0).approver1 = "Manager"
    g_rules(0).threshold2 = 500000
    g_rules(0).approver2 = "Director"
    g_rules(0).threshold3 = 2000000
    g_rules(0).approver3 = "VP"
    g_rules(0).defaultApprover = "CFO"

    ' Marketing部
    g_rules(1).department = "Marketing"
    g_rules(1).threshold1 = 30000
    g_rules(1).approver1 = "Manager"
    g_rules(1).threshold2 = 300000
    g_rules(1).approver2 = "Director"
    g_rules(1).threshold3 = 1500000
    g_rules(1).approver3 = "VP"
    g_rules(1).defaultApprover = "CFO"

    ' IT部
    g_rules(2).department = "IT"
    g_rules(2).threshold1 = 100000
    g_rules(2).approver1 = "Manager"
    g_rules(2).threshold2 = 800000
    g_rules(2).approver2 = "Director"
    g_rules(2).threshold3 = 3000000
    g_rules(2).approver3 = "VP"
    g_rules(2).defaultApprover = "CFO"

    ' HR部
    g_rules(3).department = "HR"
    g_rules(3).threshold1 = 20000
    g_rules(3).approver1 = "Manager"
    g_rules(3).threshold2 = 200000
    g_rules(3).approver2 = "Director"
    g_rules(3).threshold3 = 1000000
    g_rules(3).approver3 = "VP"
    g_rules(3).defaultApprover = "CFO"

    ' Finance部
    g_rules(4).department = "Finance"
    g_rules(4).threshold1 = 10000
    g_rules(4).approver1 = "Manager"
    g_rules(4).threshold2 = 100000
    g_rules(4).approver2 = "Director"
    g_rules(4).threshold3 = 500000
    g_rules(4).approver3 = "VP"
    g_rules(4).defaultApprover = "CFO"
End Sub

' 承認権者を決定（テーブル参照のみ）
Function GetApprover(amount As Long, department As String) As String
    Dim i As Integer
    Dim rule As ApprovalRule

    For i = LBound(g_rules) To UBound(g_rules)
        rule = g_rules(i)
        If rule.department = department Then
            ' テーブルのしきい値に基づいて決定
            If amount < rule.threshold1 Then
                GetApprover = rule.approver1
            ElseIf amount < rule.threshold2 Then
                GetApprover = rule.approver2
            ElseIf amount < rule.threshold3 Then
                GetApprover = rule.approver3
            Else
                GetApprover = rule.defaultApprover
            End If
            Exit Function
        End If
    Next i

    GetApprover = "Unknown"
End Function

' 実行例
Sub TestApprovalAfter()
    Call InitializeApprovalRules

    Debug.Print "営業部 100,000円: " & GetApprover(100000, "Sales")      ' → Director
    Debug.Print "営業部 1,000,000円: " & GetApprover(1000000, "Sales")    ' → VP
    Debug.Print "マーケティング部 50,000円: " & GetApprover(50000, "Marketing")  ' → Director
    Debug.Print "IT部 500,000円: " & GetApprover(500000, "IT")            ' → Director
    Debug.Print "HR部 150,000円: " & GetApprover(150000, "HR")            ' → Director
    Debug.Print "Finance部 50,000円: " & GetApprover(50000, "Finance")    ' → Director
End Sub
