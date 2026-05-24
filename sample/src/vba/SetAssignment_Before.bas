' SetAssignment_Before.bas
' Set xx = Function() パターンを含む if-else チェーンの検出テスト。
'
' 各ブランチの葉ノードが Set result = SomeFactory() の形式。
' 修正前: analyzeAssignments が IfStatement を再帰走査し
'         CallExpression を検出 → complexity 100 / Complex logic と誤判定。
' 修正後: IfStatement ノードをスキップ → complexity 0 / Simple assignments。

Function GetHandlerByDepartment(department As String, amount As Long) As Object
    If department = "Engineering" Then
        If amount < 100000 Then
            Set GetHandlerByDepartment = CreateTeamLeadHandler()
        ElseIf amount < 500000 Then
            Set GetHandlerByDepartment = CreateManagerHandler()
        Else
            Set GetHandlerByDepartment = CreateDirectorHandler()
        End If
    ElseIf department = "Marketing" Then
        If amount < 80000 Then
            Set GetHandlerByDepartment = CreateTeamLeadHandler()
        ElseIf amount < 400000 Then
            Set GetHandlerByDepartment = CreateManagerHandler()
        Else
            Set GetHandlerByDepartment = CreateDirectorHandler()
        End If
    ElseIf department = "Finance" Then
        If amount < 60000 Then
            Set GetHandlerByDepartment = CreateTeamLeadHandler()
        ElseIf amount < 350000 Then
            Set GetHandlerByDepartment = CreateManagerHandler()
        Else
            Set GetHandlerByDepartment = CreateDirectorHandler()
        End If
    End If
End Function
