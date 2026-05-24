' DetectorLimits.bas
' テーブル駆動検出器の限界を段階的に示すサンプル。
' 各ケースに対応する T-0x タグを付けている。

' ===========================================================================
' T-01: 条件変数の代入形状が均一なのに、形状情報が出力されない
'
' 問題: department / amount はどちらも Sheet1.Cells(row, col).Value という
'       同一 AST 形状で読み込まれている。
'       現在の出力は assignedLines（行番号）のみで、形状は不明。
'       本来は「代入部分もテーブル駆動化できる」と提案できるはず。
'
' 期待する将来の出力:
'   conditionVariables[0].assignmentAstShape = "IDENT=MEMBER(MEMBER(IDENT,CALL(NUM)),IDENT)"
'   conditionVariables[1].assignmentAstShape = "IDENT=MEMBER(MEMBER(IDENT,CALL(NUM)),IDENT)"
'   hasUniformAssignment = true
' ===========================================================================
Function GetApproverFromSheet(row As Long) As String
    Dim department As String
    Dim amount As Long

    department = Sheet1.Cells(row, 1).Value
    amount = Sheet1.Cells(row, 3).Value

    If department = "Sales" Then
        If amount < 50000 Then
            GetApproverFromSheet = "Manager"
        ElseIf amount < 500000 Then
            GetApproverFromSheet = "Director"
        Else
            GetApproverFromSheet = "VP"
        End If
    ElseIf department = "IT" Then
        If amount < 100000 Then
            GetApproverFromSheet = "Manager"
        ElseIf amount < 800000 Then
            GetApproverFromSheet = "Director"
        Else
            GetApproverFromSheet = "VP"
        End If
    ElseIf department = "HR" Then
        If amount < 20000 Then
            GetApproverFromSheet = "Manager"
        ElseIf amount < 200000 Then
            GetApproverFromSheet = "Director"
        Else
            GetApproverFromSheet = "VP"
        End If
    End If
End Function

' ===========================================================================
' T-02: 多段ネストの条件変数がフラットに収集され、レベル情報が失われる
'
' 問題: department (Level 0)・budgetType (Level 1)・amount (Level 2) の
'       3変数がすべて同列に conditionVariables へ並ぶ。
'       どれが外側キーでどれがしきい値かが出力から判断できない。
'
' 期待する将来の出力:
'   conditionVariables[0] = { name: 'department', level: 0, ... }
'   conditionVariables[1] = { name: 'budgetType',  level: 1, ... }
'   conditionVariables[2] = { name: 'amount',      level: 2, ... }  ← しきい値
' ===========================================================================
Function GetApproverByBudget(department As String, budgetType As String, amount As Long) As String
    If department = "Engineering" Then
        If budgetType = "Capex" Then
            If amount < 100000 Then
                GetApproverByBudget = "TeamLead"
            ElseIf amount < 500000 Then
                GetApproverByBudget = "Manager"
            Else
                GetApproverByBudget = "Director"
            End If
        ElseIf budgetType = "Opex" Then
            If amount < 50000 Then
                GetApproverByBudget = "TeamLead"
            ElseIf amount < 300000 Then
                GetApproverByBudget = "Manager"
            Else
                GetApproverByBudget = "Director"
            End If
        End If
    ElseIf department = "Marketing" Then
        If budgetType = "Capex" Then
            If amount < 80000 Then
                GetApproverByBudget = "TeamLead"
            ElseIf amount < 400000 Then
                GetApproverByBudget = "Manager"
            Else
                GetApproverByBudget = "Director"
            End If
        ElseIf budgetType = "Opex" Then
            If amount < 30000 Then
                GetApproverByBudget = "TeamLead"
            ElseIf amount < 200000 Then
                GetApproverByBudget = "Manager"
            Else
                GetApproverByBudget = "Director"
            End If
        End If
    ElseIf department = "Finance" Then
        If budgetType = "Capex" Then
            If amount < 60000 Then
                GetApproverByBudget = "TeamLead"
            ElseIf amount < 350000 Then
                GetApproverByBudget = "Manager"
            Else
                GetApproverByBudget = "Director"
            End If
        ElseIf budgetType = "Opex" Then
            If amount < 25000 Then
                GetApproverByBudget = "TeamLead"
            ElseIf amount < 150000 Then
                GetApproverByBudget = "Manager"
            Else
                GetApproverByBudget = "Director"
            End If
        End If
    End If
End Function

' ===========================================================================
' T-03: 条件式が IDENT = STR 以外の形式のとき、キー変数の抽出が壊れる
'
' ケース A: メンバーアクセス条件  task.Category = "Infrastructure"
'   → 現状: conditionVariables に "task" が入るか、または何も入らない
'   → 期待: キー式 = "task.Category"、値 = ["Infrastructure", "Application", "Security"]
'
' ケース B: 関数呼び出し条件  GetPriority(taskType) = "High"
'   → 現状: conditionVariables に "tasktype" が入るか検出失敗
'   → 期待: キー式 = "GetPriority(taskType)"、値 = ["High", "Medium", "Low"]
'
' 形状比較ファースト（T-03 案）ならば条件 AST 形状が全ブランチで
' MEMBER(IDENT,IDENT)=STR または CALL(IDENT)=STR と一致することを先に確認し、
' 差分（リテラル値）をキー値として抽出できる。
' ===========================================================================

' ケース A: メンバーアクセス
Function GetTaskApprover(task As Object, amount As Long) As String
    If task.Category = "Infrastructure" Then
        If amount < 100000 Then
            GetTaskApprover = "Manager"
        ElseIf amount < 500000 Then
            GetTaskApprover = "Director"
        Else
            GetTaskApprover = "VP"
        End If
    ElseIf task.Category = "Application" Then
        If amount < 80000 Then
            GetTaskApprover = "Manager"
        ElseIf amount < 400000 Then
            GetTaskApprover = "Director"
        Else
            GetTaskApprover = "VP"
        End If
    ElseIf task.Category = "Security" Then
        If amount < 50000 Then
            GetTaskApprover = "Manager"
        ElseIf amount < 300000 Then
            GetTaskApprover = "Director"
        Else
            GetTaskApprover = "VP"
        End If
    End If
End Function

' ケース B: 関数呼び出し
Function GetRequestApprover(taskType As String, amount As Long) As String
    If GetPriority(taskType) = "High" Then
        If amount < 50000 Then
            GetRequestApprover = "Manager"
        ElseIf amount < 300000 Then
            GetRequestApprover = "Director"
        Else
            GetRequestApprover = "VP"
        End If
    ElseIf GetPriority(taskType) = "Medium" Then
        If amount < 30000 Then
            GetRequestApprover = "Manager"
        ElseIf amount < 200000 Then
            GetRequestApprover = "Director"
        Else
            GetRequestApprover = "VP"
        End If
    ElseIf GetPriority(taskType) = "Low" Then
        If amount < 10000 Then
            GetRequestApprover = "Manager"
        ElseIf amount < 100000 Then
            GetRequestApprover = "Director"
        Else
            GetRequestApprover = "VP"
        End If
    End If
End Function

' ===========================================================================
' T-04: 各レベルの条件形状の均一性チェック
'
' 問題: 現在は outerChain の最初のブランチの内側条件だけを見て
'       isNumericThresholdCondition() で判定している。
'       内側条件の形状がブランチによって異なる場合に検出が不正確になる。
'
' 例: Sales/HR は内側が amount < NUM（しきい値）
'     IT は内側が IsHighPriority(cond) = Bool（関数呼び出し - 形状違い）
'
' 期待する将来の出力:
'   levelShapes[0] = { level: 0, shape: "(IDENT=STR)", isUniform: true }
'   levelShapes[1] = { level: 1, isUniform: false,
'                       warning: "IT ブランチの内側形状が他と異なる" }
' ===========================================================================
Function GetApproverMixed(department As String, amount As Long, condition As String) As String
    If department = "Sales" Then
        If amount < 50000 Then
            GetApproverMixed = "Manager"
        ElseIf amount < 500000 Then
            GetApproverMixed = "Director"
        Else
            GetApproverMixed = "VP"
        End If
    ElseIf department = "IT" Then
        If IsHighPriority(condition) Then
            GetApproverMixed = "Director"
        ElseIf IsLowPriority(condition) Then
            GetApproverMixed = "Manager"
        Else
            GetApproverMixed = "Staff"
        End If
    ElseIf department = "HR" Then
        If amount < 20000 Then
            GetApproverMixed = "Manager"
        ElseIf amount < 200000 Then
            GetApproverMixed = "Director"
        Else
            GetApproverMixed = "VP"
        End If
    End If
End Function
