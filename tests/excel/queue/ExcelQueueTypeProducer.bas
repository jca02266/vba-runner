Attribute VB_Name = "ExcelQueueTypeProducer"
Option Explicit

Public Type RecordT
    Id As Long
End Type

Public Function MakeRecords() As RecordT()
    Dim values(0 To 0) As RecordT
    values(0).Id = 9
    MakeRecords = values
End Function
