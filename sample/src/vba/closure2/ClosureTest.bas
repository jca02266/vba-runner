Attribute VB_Name = "ClosureTest"
Option Explicit

' テスト用レコードを作って RunKeyBreak を動かし、結果を文字列で返す
Function RunSubtotalTest() As String
    Dim records As New Collection

    Dim r As SalesRecord

    Set r = New SalesRecord
    r.Department = "営業"
    r.Amount = 100
    records.Add r

    Set r = New SalesRecord
    r.Department = "営業"
    r.Amount = 200
    records.Add r

    Set r = New SalesRecord
    r.Department = "開発"
    r.Amount = 150
    records.Add r

    Set r = New SalesRecord
    r.Department = "開発"
    r.Amount = 300
    records.Add r

    Set r = New SalesRecord
    r.Department = "総務"
    r.Amount = 80
    records.Add r

    Dim body As New DeptSubtotalBody
    RunKeyBreak records, body
    RunSubtotalTest = body.GetResult()
End Function

Function RunCountTest() As String
    Dim records As New Collection

    Dim r As SalesRecord

    Set r = New SalesRecord : r.Department = "営業" : r.Amount = 100 : records.Add r
    Set r = New SalesRecord : r.Department = "営業" : r.Amount = 200 : records.Add r
    Set r = New SalesRecord : r.Department = "開発" : r.Amount = 150 : records.Add r
    Set r = New SalesRecord : r.Department = "開発" : r.Amount = 300 : records.Add r
    Set r = New SalesRecord : r.Department = "総務" : r.Amount = 80  : records.Add r

    Dim body As New DeptCountBody
    RunKeyBreak records, body
    RunCountTest = body.GetResult()
End Function
