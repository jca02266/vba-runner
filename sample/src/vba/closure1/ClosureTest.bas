Attribute VB_Name = "ClosureTest"
Option Explicit

Function MakeRecords() As Collection
    Dim records As New Collection
    Dim r As SalesRecord

    Set r = New SalesRecord : r.Department = "営業" : r.Amount = 100 : records.Add r
    Set r = New SalesRecord : r.Department = "営業" : r.Amount = 200 : records.Add r
    Set r = New SalesRecord : r.Department = "開発" : r.Amount = 150 : records.Add r
    Set r = New SalesRecord : r.Department = "開発" : r.Amount = 300 : records.Add r
    Set r = New SalesRecord : r.Department = "総務" : r.Amount = 80  : records.Add r

    Set MakeRecords = records
End Function

Function RunSubtotalTest() As String
    Dim body As New DeptSubtotalBody
    RunKeyBreak MakeRecords(), body
    RunSubtotalTest = body.GetResult()
End Function

Function RunCountTest() As String
    Dim body As New DeptCountBody
    RunKeyBreak MakeRecords(), body
    RunCountTest = body.GetResult()
End Function
