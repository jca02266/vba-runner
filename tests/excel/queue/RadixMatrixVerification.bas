Attribute VB_Name = "RadixMatrixVerification"
Option Explicit

Private resultFile As Integer
Private Const QUEUE_SOURCE_SHA256 As String = ""

Public Sub RunRadixMatrixVerification()
    Dim values As Variant, kind As Variant, value As Variant
    Dim path As String
    path = ThisWorkbook.Path & Application.PathSeparator & "RadixMatrix.result"
    values = Array("0", "1", "-1", "&H7FFF", "&H8000", "&HFFFF", _
        "&H7FFFFFFF", "&H80000000", "&HFFFFFFFF", _
        "&H7FFFFFFFFFFFFFFF", "&H8000000000000000", "&HFFFFFFFFFFFFFFFF", _
        "&O177777", "&O37777777777", "&O1777777777777777777777", _
        "9007199254740993", "-9223372036854775808")
    resultFile = FreeFile
    Open path For Output As #resultFile
    Print #resultFile, "RADIX_MATRIX_BEGIN"
    For Each kind In Array("CByte", "CInt", "CLng", "CLngLng", "CDec", "CCur", "Val")
        For Each value In values
            EmitCase "XL-175", CStr(kind), CStr(value)
        Next value
    Next kind
    EmitImplicitWidthCase "XL-176", "&H100", "Byte"
    EmitImplicitWidthCase "XL-177", "&H10000", "Integer"
    EmitImplicitWidthCase "XL-178", "&H100000000", "Long"
    EmitImplicitWidthCase "XL-179", "&H10000000000000000", "LongLong"
    Print #resultFile, "QUEUE_SOURCE_SHA256=" & QUEUE_SOURCE_SHA256
    Print #resultFile, "RADIX_MATRIX_COMPLETE=True"
    Close #resultFile
End Sub

Private Sub EmitImplicitWidthCase(ByVal probeId As String, ByVal text As String, ByVal targetKind As String)
    Dim source As Variant, result As Variant, errNo As Long
    On Error Resume Next
    Err.Clear
    Select Case targetKind
        Case "Byte": source = &H100: Dim byteValue As Byte: byteValue = source: result = byteValue
        Case "Integer": source = &H10000: Dim integerValue As Integer: integerValue = source: result = integerValue
        Case "Long": source = &H100000000: Dim longValue As Long: longValue = source: result = longValue
        Case "LongLong": source = &H10000000000000000^: Dim longLongValue As LongLong: longLongValue = source: result = longLongValue
    End Select
    errNo = Err.Number
    If errNo <> 0 Then
        Print #resultFile, "CASE=" & probeId & " KIND=" & targetKind & " INPUT=" & text & " ERR=" & CStr(errNo)
    Else
        Print #resultFile, "CASE=" & probeId & " KIND=" & targetKind & " INPUT=" & text & " TYPE=" & TypeName(result) & " VALUE=" & CStr(result)
    End If
    Err.Clear
    On Error GoTo 0
End Sub

Private Sub EmitCase(ByVal probeId As String, ByVal kind As String, ByVal text As String)
    Dim result As Variant, errNo As Long, line As String
    On Error Resume Next
    Err.Clear
    Select Case kind
        Case "CByte": result = CByte(text)
        Case "CInt": result = CInt(text)
        Case "CLng": result = CLng(text)
        Case "CLngLng": result = CLngLng(text)
        Case "CDec": result = CDec(text)
        Case "CCur": result = CCur(text)
        Case "Val": result = Val(text)
    End Select
    errNo = Err.Number
    If errNo <> 0 Then
        line = "CASE=" & probeId & " KIND=" & kind & " INPUT=" & Encode(text) & _
            " ERR=" & CStr(errNo)
    Else
        line = "CASE=" & probeId & " KIND=" & kind & " INPUT=" & Encode(text) & _
            " TYPE=" & TypeName(result) & " VALUE=" & Encode(CStr(result))
    End If
    Print #resultFile, line
    Debug.Print line
    Err.Clear
    On Error GoTo 0
End Sub

Private Function Encode(ByVal value As String) As String
    Encode = Replace(Replace(Replace(value, vbCr, "<CR>"), vbLf, "<LF>"), vbTab, "<TAB>")
End Function
