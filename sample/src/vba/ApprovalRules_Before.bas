' ============================================
' リファクタリング前：分岐地獄パターン
' 請求書の金額と部門から、承認権者を決定する
' ============================================

Function GetApprover(amount As Long, department As String) As String
    ' 営業部
    If department = "Sales" Then
        If amount < 50000 Then
            GetApprover = "Manager"
        ElseIf amount < 500000 Then
            GetApprover = "Director"
        ElseIf amount < 2000000 Then
            GetApprover = "VP"
        Else
            GetApprover = "CFO"
        End If
    ' マーケティング部
    ElseIf department = "Marketing" Then
        If amount < 30000 Then
            GetApprover = "Manager"
        ElseIf amount < 300000 Then
            GetApprover = "Director"
        ElseIf amount < 1500000 Then
            GetApprover = "VP"
        Else
            GetApprover = "CFO"
        End If
    ' IT部
    ElseIf department = "IT" Then
        If amount < 100000 Then
            GetApprover = "Manager"
        ElseIf amount < 800000 Then
            GetApprover = "Director"
        ElseIf amount < 3000000 Then
            GetApprover = "VP"
        Else
            GetApprover = "CFO"
        End If
    ' HR部
    ElseIf department = "HR" Then
        If amount < 20000 Then
            GetApprover = "Manager"
        ElseIf amount < 200000 Then
            GetApprover = "Director"
        ElseIf amount < 1000000 Then
            GetApprover = "VP"
        Else
            GetApprover = "CFO"
        End If
    ' Finance部（最も厳しい）
    ElseIf department = "Finance" Then
        If amount < 10000 Then
            GetApprover = "Manager"
        ElseIf amount < 100000 Then
            GetApprover = "Director"
        ElseIf amount < 500000 Then
            GetApprover = "VP"
        Else
            GetApprover = "CFO"
        End If
    Else
        GetApprover = "Unknown"
    End If
End Function

' 実行例
Sub TestApprovalBefore()
    Debug.Print "営業部 100,000円: " & GetApprover(100000, "Sales")      ' → Director
    Debug.Print "営業部 1,000,000円: " & GetApprover(1000000, "Sales")    ' → VP
    Debug.Print "マーケティング部 50,000円: " & GetApprover(50000, "Marketing")  ' → Director
    Debug.Print "IT部 500,000円: " & GetApprover(500000, "IT")            ' → Director
    Debug.Print "HR部 150,000円: " & GetApprover(150000, "HR")            ' → Director
    Debug.Print "Finance部 50,000円: " & GetApprover(50000, "Finance")    ' → Director
End Sub
