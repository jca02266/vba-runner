Attribute VB_Name = "ExcelQueueVerification"
Option Explicit

Private Type MatrixRecord
    Values(0 To 1, 0 To 1) As Long
End Type

Private Type StringMatrixRecord
    Values(0 To 1, 0 To 1) As String * 2
End Type

Private Type PointRecord
    X As Long
    Y As Integer
End Type

Public Sub RunExcelQueueVerification()
    Dim root As String
    root = Environ$("TEMP") & Application.PathSeparator & "vba-runner-xl-queue"
    On Error Resume Next
    MkDir root
    On Error GoTo 0
    VerifyMatrix root & Application.PathSeparator & "XL-001.bin"
    VerifyStringMatrix root & Application.PathSeparator & "XL-002.bin"
    VerifyByteArray root & Application.PathSeparator & "XL-003.bin"
    VerifyIntegerArray root & Application.PathSeparator & "XL-004.bin"
    VerifyLongArray root & Application.PathSeparator & "XL-005.bin"
    VerifyFixedStringArray root & Application.PathSeparator & "XL-006.bin"
    VerifyPointArray root & Application.PathSeparator & "XL-007.bin"
    Debug.Print "RESULT_ROOT=" & root
End Sub

Private Sub VerifyMatrix(ByVal path As String)
    Dim r As MatrixRecord, f As Integer
    r.Values(0, 0) = 1: r.Values(0, 1) = 2: r.Values(1, 0) = 3: r.Values(1, 1) = 4
    f = FreeFile: Open path For Binary Access Write As #f: Put #f, , r: Close #f
    PrintBytes "XL-001", path
End Sub

Private Sub VerifyStringMatrix(ByVal path As String)
    Dim r As StringMatrixRecord, f As Integer
    r.Values(0, 0) = "A": r.Values(0, 1) = "B": r.Values(1, 0) = "C": r.Values(1, 1) = "D"
    f = FreeFile: Open path For Binary Access Write As #f: Put #f, , r: Close #f
    PrintBytes "XL-002", path
End Sub

Private Sub VerifyByteArray(ByVal path As String)
    Dim values(0 To 1, 0 To 1) As Byte, f As Integer
    values(0, 0) = 1: values(0, 1) = 2: values(1, 0) = 3: values(1, 1) = 4
    f = FreeFile: Open path For Binary Access Write As #f: Put #f, , values: Close #f
    PrintBytes "XL-003", path
End Sub

Private Sub VerifyIntegerArray(ByVal path As String)
    Dim values(0 To 1, 0 To 1) As Integer, f As Integer
    values(0, 0) = 1: values(0, 1) = 2: values(1, 0) = 3: values(1, 1) = 4
    f = FreeFile: Open path For Binary Access Write As #f: Put #f, , values: Close #f
    PrintBytes "XL-004", path
End Sub

Private Sub VerifyLongArray(ByVal path As String)
    Dim values(0 To 1, 0 To 1) As Long, f As Integer
    values(0, 0) = 1: values(0, 1) = 2: values(1, 0) = 3: values(1, 1) = 4
    f = FreeFile: Open path For Binary Access Write As #f: Put #f, , values: Close #f
    PrintBytes "XL-005", path
End Sub

Private Sub VerifyFixedStringArray(ByVal path As String)
    Dim values(0 To 1) As String * 2, f As Integer
    values(0) = "あ": values(1) = "A"
    f = FreeFile: Open path For Binary Access Write As #f: Put #f, , values: Close #f
    PrintBytes "XL-006", path
End Sub

Private Sub VerifyPointArray(ByVal path As String)
    Dim points(0 To 1) As PointRecord, f As Integer
    points(0).X = 1: points(0).Y = 2: points(1).X = 3: points(1).Y = 4
    f = FreeFile: Open path For Binary Access Write As #f: Put #f, , points: Close #f
    PrintBytes "XL-007", path
End Sub

Private Sub PrintBytes(ByVal id As String, ByVal path As String)
    Dim f As Integer, p As Long, n As Long, value As Byte, line As String
    n = FileLen(path): f = FreeFile: Open path For Binary Access Read As #f
    For p = 1 To n
        Get #f, p, value
        If Len(line) > 0 Then line = line & " "
        line = line & Right$("0" & Hex$(value), 2)
    Next p
    Close #f
    Debug.Print id & " LOF=" & n & " BYTES=" & line
End Sub
