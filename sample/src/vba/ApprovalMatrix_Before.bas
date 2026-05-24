' ============================================
' 承認マトリクス（Before: 4段階ネスト・複数文副作用版）
' 部門 × 予算種別 × 緊急度 × 金額 の4次元分岐
' 各葉に approvalLog 追記と approvalCount インクリメントを追加
' ============================================

Function GetApprovalLevel(department As String, budgetType As String, urgency As String, amount As Long) As String
    Dim approvalLog As String
    Dim approvalCount As Long
    If department = "Engineering" Then
        If budgetType = "Capex" Then
            If urgency = "Emergency" Then
                If amount < 100000 Then
                    GetApprovalLevel = "TeamLead"
                    approvalLog = approvalLog & "TeamLead"
                    approvalCount = approvalCount + 1
                ElseIf amount < 1000000 Then
                    GetApprovalLevel = "Manager"
                    approvalLog = approvalLog & "Manager"
                    approvalCount = approvalCount + 1
                ElseIf amount < 5000000 Then
                    GetApprovalLevel = "Director"
                    approvalLog = approvalLog & "Director"
                    approvalCount = approvalCount + 1
                Else
                    GetApprovalLevel = "VP"
                    approvalLog = approvalLog & "VP"
                    approvalCount = approvalCount + 1
                End If
            ElseIf urgency = "Routine" Then
                If amount < 50000 Then
                    GetApprovalLevel = "TeamLead"
                    approvalLog = approvalLog & "TeamLead"
                    approvalCount = approvalCount + 1
                ElseIf amount < 500000 Then
                    GetApprovalLevel = "Manager"
                    approvalLog = approvalLog & "Manager"
                    approvalCount = approvalCount + 1
                ElseIf amount < 3000000 Then
                    GetApprovalLevel = "Director"
                    approvalLog = approvalLog & "Director"
                    approvalCount = approvalCount + 1
                Else
                    GetApprovalLevel = "VP"
                    approvalLog = approvalLog & "VP"
                    approvalCount = approvalCount + 1
                End If
            End If
        ElseIf budgetType = "Opex" Then
            If urgency = "Emergency" Then
                If amount < 50000 Then
                    GetApprovalLevel = "TeamLead"
                    approvalLog = approvalLog & "TeamLead"
                    approvalCount = approvalCount + 1
                ElseIf amount < 300000 Then
                    GetApprovalLevel = "Manager"
                    approvalLog = approvalLog & "Manager"
                    approvalCount = approvalCount + 1
                ElseIf amount < 1000000 Then
                    GetApprovalLevel = "Director"
                    approvalLog = approvalLog & "Director"
                    approvalCount = approvalCount + 1
                Else
                    GetApprovalLevel = "VP"
                    approvalLog = approvalLog & "VP"
                    approvalCount = approvalCount + 1
                End If
            ElseIf urgency = "Routine" Then
                If amount < 30000 Then
                    GetApprovalLevel = "TeamLead"
                    approvalLog = approvalLog & "TeamLead"
                    approvalCount = approvalCount + 1
                ElseIf amount < 200000 Then
                    GetApprovalLevel = "Manager"
                    approvalLog = approvalLog & "Manager"
                    approvalCount = approvalCount + 1
                ElseIf amount < 800000 Then
                    GetApprovalLevel = "Director"
                    approvalLog = approvalLog & "Director"
                    approvalCount = approvalCount + 1
                Else
                    GetApprovalLevel = "VP"
                    approvalLog = approvalLog & "VP"
                    approvalCount = approvalCount + 1
                End If
            End If
        ElseIf budgetType = "Travel" Then
            If urgency = "Emergency" Then
                If amount < 30000 Then
                    GetApprovalLevel = "Self"
                    approvalLog = approvalLog & "Self"
                    approvalCount = approvalCount + 1
                ElseIf amount < 100000 Then
                    GetApprovalLevel = "TeamLead"
                    approvalLog = approvalLog & "TeamLead"
                    approvalCount = approvalCount + 1
                ElseIf amount < 300000 Then
                    GetApprovalLevel = "Manager"
                    approvalLog = approvalLog & "Manager"
                    approvalCount = approvalCount + 1
                Else
                    GetApprovalLevel = "Director"
                    approvalLog = approvalLog & "Director"
                    approvalCount = approvalCount + 1
                End If
            ElseIf urgency = "Routine" Then
                If amount < 20000 Then
                    GetApprovalLevel = "Self"
                    approvalLog = approvalLog & "Self"
                    approvalCount = approvalCount + 1
                ElseIf amount < 80000 Then
                    GetApprovalLevel = "TeamLead"
                    approvalLog = approvalLog & "TeamLead"
                    approvalCount = approvalCount + 1
                ElseIf amount < 200000 Then
                    GetApprovalLevel = "Manager"
                    approvalLog = approvalLog & "Manager"
                    approvalCount = approvalCount + 1
                Else
                    GetApprovalLevel = "Director"
                    approvalLog = approvalLog & "Director"
                    approvalCount = approvalCount + 1
                End If
            End If
        End If
    ElseIf department = "Marketing" Then
        If budgetType = "Capex" Then
            If urgency = "Emergency" Then
                If amount < 80000 Then
                    GetApprovalLevel = "TeamLead"
                    approvalLog = approvalLog & "TeamLead"
                    approvalCount = approvalCount + 1
                ElseIf amount < 800000 Then
                    GetApprovalLevel = "Manager"
                    approvalLog = approvalLog & "Manager"
                    approvalCount = approvalCount + 1
                ElseIf amount < 4000000 Then
                    GetApprovalLevel = "Director"
                    approvalLog = approvalLog & "Director"
                    approvalCount = approvalCount + 1
                Else
                    GetApprovalLevel = "VP"
                    approvalLog = approvalLog & "VP"
                    approvalCount = approvalCount + 1
                End If
            ElseIf urgency = "Routine" Then
                If amount < 40000 Then
                    GetApprovalLevel = "TeamLead"
                    approvalLog = approvalLog & "TeamLead"
                    approvalCount = approvalCount + 1
                ElseIf amount < 400000 Then
                    GetApprovalLevel = "Manager"
                    approvalLog = approvalLog & "Manager"
                    approvalCount = approvalCount + 1
                ElseIf amount < 2000000 Then
                    GetApprovalLevel = "Director"
                    approvalLog = approvalLog & "Director"
                    approvalCount = approvalCount + 1
                Else
                    GetApprovalLevel = "VP"
                    approvalLog = approvalLog & "VP"
                    approvalCount = approvalCount + 1
                End If
            End If
        ElseIf budgetType = "Opex" Then
            If urgency = "Emergency" Then
                If amount < 40000 Then
                    GetApprovalLevel = "TeamLead"
                    approvalLog = approvalLog & "TeamLead"
                    approvalCount = approvalCount + 1
                ElseIf amount < 250000 Then
                    GetApprovalLevel = "Manager"
                    approvalLog = approvalLog & "Manager"
                    approvalCount = approvalCount + 1
                ElseIf amount < 800000 Then
                    GetApprovalLevel = "Director"
                    approvalLog = approvalLog & "Director"
                    approvalCount = approvalCount + 1
                Else
                    GetApprovalLevel = "VP"
                    approvalLog = approvalLog & "VP"
                    approvalCount = approvalCount + 1
                End If
            ElseIf urgency = "Routine" Then
                If amount < 25000 Then
                    GetApprovalLevel = "TeamLead"
                    approvalLog = approvalLog & "TeamLead"
                    approvalCount = approvalCount + 1
                ElseIf amount < 150000 Then
                    GetApprovalLevel = "Manager"
                    approvalLog = approvalLog & "Manager"
                    approvalCount = approvalCount + 1
                ElseIf amount < 600000 Then
                    GetApprovalLevel = "Director"
                    approvalLog = approvalLog & "Director"
                    approvalCount = approvalCount + 1
                Else
                    GetApprovalLevel = "VP"
                    approvalLog = approvalLog & "VP"
                    approvalCount = approvalCount + 1
                End If
            End If
        ElseIf budgetType = "Travel" Then
            If urgency = "Emergency" Then
                If amount < 25000 Then
                    GetApprovalLevel = "Self"
                    approvalLog = approvalLog & "Self"
                    approvalCount = approvalCount + 1
                ElseIf amount < 80000 Then
                    GetApprovalLevel = "TeamLead"
                    approvalLog = approvalLog & "TeamLead"
                    approvalCount = approvalCount + 1
                ElseIf amount < 250000 Then
                    GetApprovalLevel = "Manager"
                    approvalLog = approvalLog & "Manager"
                    approvalCount = approvalCount + 1
                Else
                    GetApprovalLevel = "Director"
                    approvalLog = approvalLog & "Director"
                    approvalCount = approvalCount + 1
                End If
            ElseIf urgency = "Routine" Then
                If amount < 15000 Then
                    GetApprovalLevel = "Self"
                    approvalLog = approvalLog & "Self"
                    approvalCount = approvalCount + 1
                ElseIf amount < 60000 Then
                    GetApprovalLevel = "TeamLead"
                    approvalLog = approvalLog & "TeamLead"
                    approvalCount = approvalCount + 1
                ElseIf amount < 150000 Then
                    GetApprovalLevel = "Manager"
                    approvalLog = approvalLog & "Manager"
                    approvalCount = approvalCount + 1
                Else
                    GetApprovalLevel = "Director"
                    approvalLog = approvalLog & "Director"
                    approvalCount = approvalCount + 1
                End If
            End If
        End If
    ElseIf department = "Finance" Then
        If budgetType = "Capex" Then
            If urgency = "Emergency" Then
                If amount < 200000 Then
                    GetApprovalLevel = "TeamLead"
                    approvalLog = approvalLog & "TeamLead"
                    approvalCount = approvalCount + 1
                ElseIf amount < 2000000 Then
                    GetApprovalLevel = "Manager"
                    approvalLog = approvalLog & "Manager"
                    approvalCount = approvalCount + 1
                ElseIf amount < 10000000 Then
                    GetApprovalLevel = "Director"
                    approvalLog = approvalLog & "Director"
                    approvalCount = approvalCount + 1
                Else
                    GetApprovalLevel = "CFO"
                    approvalLog = approvalLog & "CFO"
                    approvalCount = approvalCount + 1
                End If
            ElseIf urgency = "Routine" Then
                If amount < 100000 Then
                    GetApprovalLevel = "TeamLead"
                    approvalLog = approvalLog & "TeamLead"
                    approvalCount = approvalCount + 1
                ElseIf amount < 1000000 Then
                    GetApprovalLevel = "Manager"
                    approvalLog = approvalLog & "Manager"
                    approvalCount = approvalCount + 1
                ElseIf amount < 5000000 Then
                    GetApprovalLevel = "Director"
                    approvalLog = approvalLog & "Director"
                    approvalCount = approvalCount + 1
                Else
                    GetApprovalLevel = "CFO"
                    approvalLog = approvalLog & "CFO"
                    approvalCount = approvalCount + 1
                End If
            End If
        ElseIf budgetType = "Opex" Then
            If urgency = "Emergency" Then
                If amount < 100000 Then
                    GetApprovalLevel = "TeamLead"
                    approvalLog = approvalLog & "TeamLead"
                    approvalCount = approvalCount + 1
                ElseIf amount < 600000 Then
                    GetApprovalLevel = "Manager"
                    approvalLog = approvalLog & "Manager"
                    approvalCount = approvalCount + 1
                ElseIf amount < 2000000 Then
                    GetApprovalLevel = "Director"
                    approvalLog = approvalLog & "Director"
                    approvalCount = approvalCount + 1
                Else
                    GetApprovalLevel = "CFO"
                    approvalLog = approvalLog & "CFO"
                    approvalCount = approvalCount + 1
                End If
            ElseIf urgency = "Routine" Then
                If amount < 60000 Then
                    GetApprovalLevel = "TeamLead"
                    approvalLog = approvalLog & "TeamLead"
                    approvalCount = approvalCount + 1
                ElseIf amount < 400000 Then
                    GetApprovalLevel = "Manager"
                    approvalLog = approvalLog & "Manager"
                    approvalCount = approvalCount + 1
                ElseIf amount < 1500000 Then
                    GetApprovalLevel = "Director"
                    approvalLog = approvalLog & "Director"
                    approvalCount = approvalCount + 1
                Else
                    GetApprovalLevel = "CFO"
                    approvalLog = approvalLog & "CFO"
                    approvalCount = approvalCount + 1
                End If
            End If
        ElseIf budgetType = "Travel" Then
            If urgency = "Emergency" Then
                If amount < 50000 Then
                    GetApprovalLevel = "Self"
                    approvalLog = approvalLog & "Self"
                    approvalCount = approvalCount + 1
                ElseIf amount < 200000 Then
                    GetApprovalLevel = "TeamLead"
                    approvalLog = approvalLog & "TeamLead"
                    approvalCount = approvalCount + 1
                ElseIf amount < 500000 Then
                    GetApprovalLevel = "Manager"
                    approvalLog = approvalLog & "Manager"
                    approvalCount = approvalCount + 1
                Else
                    GetApprovalLevel = "Director"
                    approvalLog = approvalLog & "Director"
                    approvalCount = approvalCount + 1
                End If
            ElseIf urgency = "Routine" Then
                If amount < 30000 Then
                    GetApprovalLevel = "Self"
                    approvalLog = approvalLog & "Self"
                    approvalCount = approvalCount + 1
                ElseIf amount < 100000 Then
                    GetApprovalLevel = "TeamLead"
                    approvalLog = approvalLog & "TeamLead"
                    approvalCount = approvalCount + 1
                ElseIf amount < 300000 Then
                    GetApprovalLevel = "Manager"
                    approvalLog = approvalLog & "Manager"
                    approvalCount = approvalCount + 1
                Else
                    GetApprovalLevel = "Director"
                    approvalLog = approvalLog & "Director"
                    approvalCount = approvalCount + 1
                End If
            End If
        End If
    ElseIf department = "Operations" Then
        If budgetType = "Capex" Then
            If urgency = "Emergency" Then
                If amount < 150000 Then
                    GetApprovalLevel = "TeamLead"
                    approvalLog = approvalLog & "TeamLead"
                    approvalCount = approvalCount + 1
                ElseIf amount < 1500000 Then
                    GetApprovalLevel = "Manager"
                    approvalLog = approvalLog & "Manager"
                    approvalCount = approvalCount + 1
                ElseIf amount < 8000000 Then
                    GetApprovalLevel = "Director"
                    approvalLog = approvalLog & "Director"
                    approvalCount = approvalCount + 1
                Else
                    GetApprovalLevel = "VP"
                    approvalLog = approvalLog & "VP"
                    approvalCount = approvalCount + 1
                End If
            ElseIf urgency = "Routine" Then
                If amount < 80000 Then
                    GetApprovalLevel = "TeamLead"
                    approvalLog = approvalLog & "TeamLead"
                    approvalCount = approvalCount + 1
                ElseIf amount < 800000 Then
                    GetApprovalLevel = "Manager"
                    approvalLog = approvalLog & "Manager"
                    approvalCount = approvalCount + 1
                ElseIf amount < 4000000 Then
                    GetApprovalLevel = "Director"
                    approvalLog = approvalLog & "Director"
                    approvalCount = approvalCount + 1
                Else
                    GetApprovalLevel = "VP"
                    approvalLog = approvalLog & "VP"
                    approvalCount = approvalCount + 1
                End If
            End If
        ElseIf budgetType = "Opex" Then
            If urgency = "Emergency" Then
                If amount < 80000 Then
                    GetApprovalLevel = "TeamLead"
                    approvalLog = approvalLog & "TeamLead"
                    approvalCount = approvalCount + 1
                ElseIf amount < 500000 Then
                    GetApprovalLevel = "Manager"
                    approvalLog = approvalLog & "Manager"
                    approvalCount = approvalCount + 1
                ElseIf amount < 1500000 Then
                    GetApprovalLevel = "Director"
                    approvalLog = approvalLog & "Director"
                    approvalCount = approvalCount + 1
                Else
                    GetApprovalLevel = "VP"
                    approvalLog = approvalLog & "VP"
                    approvalCount = approvalCount + 1
                End If
            ElseIf urgency = "Routine" Then
                If amount < 50000 Then
                    GetApprovalLevel = "TeamLead"
                    approvalLog = approvalLog & "TeamLead"
                    approvalCount = approvalCount + 1
                ElseIf amount < 300000 Then
                    GetApprovalLevel = "Manager"
                    approvalLog = approvalLog & "Manager"
                    approvalCount = approvalCount + 1
                ElseIf amount < 1000000 Then
                    GetApprovalLevel = "Director"
                    approvalLog = approvalLog & "Director"
                    approvalCount = approvalCount + 1
                Else
                    GetApprovalLevel = "VP"
                    approvalLog = approvalLog & "VP"
                    approvalCount = approvalCount + 1
                End If
            End If
        ElseIf budgetType = "Travel" Then
            If urgency = "Emergency" Then
                If amount < 40000 Then
                    GetApprovalLevel = "Self"
                    approvalLog = approvalLog & "Self"
                    approvalCount = approvalCount + 1
                ElseIf amount < 150000 Then
                    GetApprovalLevel = "TeamLead"
                    approvalLog = approvalLog & "TeamLead"
                    approvalCount = approvalCount + 1
                ElseIf amount < 400000 Then
                    GetApprovalLevel = "Manager"
                    approvalLog = approvalLog & "Manager"
                    approvalCount = approvalCount + 1
                Else
                    GetApprovalLevel = "Director"
                    approvalLog = approvalLog & "Director"
                    approvalCount = approvalCount + 1
                End If
            ElseIf urgency = "Routine" Then
                If amount < 25000 Then
                    GetApprovalLevel = "Self"
                    approvalLog = approvalLog & "Self"
                    approvalCount = approvalCount + 1
                ElseIf amount < 90000 Then
                    GetApprovalLevel = "TeamLead"
                    approvalLog = approvalLog & "TeamLead"
                    approvalCount = approvalCount + 1
                ElseIf amount < 250000 Then
                    GetApprovalLevel = "Manager"
                    approvalLog = approvalLog & "Manager"
                    approvalCount = approvalCount + 1
                Else
                    GetApprovalLevel = "Director"
                    approvalLog = approvalLog & "Director"
                    approvalCount = approvalCount + 1
                End If
            End If
        End If
    End If
End Function
