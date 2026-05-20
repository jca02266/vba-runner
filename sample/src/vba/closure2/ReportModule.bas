Attribute VB_Name = "ReportModule"
Option Explicit

' 部門別売上小計レポート
Sub PrintDeptReport(records As Collection)
    Dim body As New DeptSubtotalBody
    RunKeyBreak records, body
End Sub

' 部門別レコード件数レポート
Sub PrintDeptCount(records As Collection)
    Dim body As New DeptCountBody
    RunKeyBreak records, body
End Sub
