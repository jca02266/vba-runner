' ============================================
' さらに発展：外部データ駆動パターン
' ルールを外部データ（辞書構造）で管理
' ============================================

' 部門ごとの承認閾値マップ
Type ApprovalLevel
    threshold As Long
    approver As String
End Type

Type DepartmentRules
    name As String
    levels() As ApprovalLevel
    defaultApprover As String
End Type

' Dictionaryを使った実装例（VBA Runtime Library 参照が必要）
Function BuildApprovalRulesDictionary() As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    ' Sales部のルール
    Dim salesRules As Object
    Set salesRules = CreateObject("Scripting.Dictionary")
    salesRules.Add "levels", Array( _
        CreateLevelTuple(50000, "Manager"), _
        CreateLevelTuple(500000, "Director"), _
        CreateLevelTuple(2000000, "VP"))
    salesRules.Add "default", "CFO"
    dict.Add "Sales", salesRules

    ' Marketing部のルール
    Dim marketingRules As Object
    Set marketingRules = CreateObject("Scripting.Dictionary")
    marketingRules.Add "levels", Array( _
        CreateLevelTuple(30000, "Manager"), _
        CreateLevelTuple(300000, "Director"), _
        CreateLevelTuple(1500000, "VP"))
    marketingRules.Add "default", "CFO"
    dict.Add "Marketing", marketingRules

    ' IT部のルール
    Dim itRules As Object
    Set itRules = CreateObject("Scripting.Dictionary")
    itRules.Add "levels", Array( _
        CreateLevelTuple(100000, "Manager"), _
        CreateLevelTuple(800000, "Director"), _
        CreateLevelTuple(3000000, "VP"))
    itRules.Add "default", "CFO"
    dict.Add "IT", itRules

    ' HR部のルール
    Dim hrRules As Object
    Set hrRules = CreateObject("Scripting.Dictionary")
    hrRules.Add "levels", Array( _
        CreateLevelTuple(20000, "Manager"), _
        CreateLevelTuple(200000, "Director"), _
        CreateLevelTuple(1000000, "VP"))
    hrRules.Add "default", "CFO"
    dict.Add "HR", hrRules

    ' Finance部のルール
    Dim financeRules As Object
    Set financeRules = CreateObject("Scripting.Dictionary")
    financeRules.Add "levels", Array( _
        CreateLevelTuple(10000, "Manager"), _
        CreateLevelTuple(100000, "Director"), _
        CreateLevelTuple(500000, "VP"))
    financeRules.Add "default", "CFO"
    dict.Add "Finance", financeRules

    Set BuildApprovalRulesDictionary = dict
End Function

' ヘルパー：(threshold, approver) タプルを作成
Function CreateLevelTuple(threshold As Long, approver As String) As Object
    Dim tuple As Object
    Set tuple = CreateObject("Scripting.Dictionary")
    tuple.Add "threshold", threshold
    tuple.Add "approver", approver
    Set CreateLevelTuple = tuple
End Function

' テーブル駆動：承認権者を決定
Function GetApproverFromTable(amount As Long, department As String, rules As Object) As String
    If Not rules.Exists(department) Then
        GetApproverFromTable = "Unknown"
        Exit Function
    End If

    Dim deptRules As Object
    Set deptRules = rules(department)

    Dim levels As Variant
    levels = deptRules("levels")

    Dim i As Integer
    For i = LBound(levels) To UBound(levels)
        If amount < levels(i)("threshold") Then
            GetApproverFromTable = levels(i)("approver")
            Exit Function
        End If
    Next i

    GetApproverFromTable = deptRules("default")
End Function

' 実行例
Sub TestApprovalAdvanced()
    Dim rules As Object
    Set rules = BuildApprovalRulesDictionary

    Debug.Print "営業部 100,000円: " & GetApproverFromTable(100000, "Sales", rules)
    Debug.Print "営業部 1,000,000円: " & GetApproverFromTable(1000000, "Sales", rules)
    Debug.Print "マーケティング部 50,000円: " & GetApproverFromTable(50000, "Marketing", rules)
    Debug.Print "IT部 500,000円: " & GetApproverFromTable(500000, "IT", rules)
    Debug.Print "HR部 150,000円: " & GetApproverFromTable(150000, "HR", rules)
    Debug.Print "Finance部 50,000円: " & GetApproverFromTable(50000, "Finance", rules)
End Sub
