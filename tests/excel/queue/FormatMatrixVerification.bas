Attribute VB_Name = "FormatMatrixVerification"
Option Explicit

Private resultFile As Integer

Public Sub RunFormatMatrixVerification()
    Dim path As String
    path = ThisWorkbook.Path & Application.PathSeparator & "FormatMatrix.result"
    resultFile = FreeFile
    Open path For Output As #resultFile
    Print #resultFile, "FORMAT_MATRIX_BEGIN"
    EmitSingleCharacterMatrix
    EmitCompoundTokenMatrix
    EmitTokenizerCornerMatrix
    Print #resultFile, "FORMAT_MATRIX_COMPLETE=True"
    Close #resultFile
End Sub

Private Sub EmitTokenizerCornerMatrix()
    Dim patterns As Variant, i As Long
    ' These patterns are extracted from splitFormatSections, the escape and
    ' quote scanners, and the numeric/date/string token dispatchers. Keep
    ' them in this dedicated matrix so normal Excel queue output stays small.
    patterns = Array( _
        "0\;0", "0"";""0", "0\", "0""literal", _
        "[abc", "\[abc\]", """[abc]""", "[Red]0", "[<=1000]0", _
        "0.00e-00", "%", "0%", ".", ",", "!@@@", "@@@", "\@", _
        """@""", ">@@@", "<@@@", "hh:mm", "h:m", "yyyy-mm-dd", _
        "m", "n", "E", "G", "ggee", "\Y""Year""yyyy" _
    )
    For i = LBound(patterns) To UBound(patterns)
        EmitArguments "TOKENIZER-" & Format$(i + 1, "00"), CStr(patterns(i))
    Next i
End Sub

Private Sub EmitSingleCharacterMatrix()
    Dim code As Long, pattern As String
    For code = 32 To 126
        pattern = Chr$(code)
        EmitArguments "ASCII-" & CStr(code), pattern
    Next code
End Sub

Private Sub EmitCompoundTokenMatrix()
    Dim patterns As Variant, i As Long
    patterns = Array( _
        "0;0", "0;[Red](0)", "Lo;Hi", "Lo0;Hi", "Lo;Hi0", _
        "[abc]", "[123]", "[<=1000]", "[Lo]", "[Lo", "Lo]", _
        "\[<=1000\]", """[<=1000]""", "[Lo]0", "0;", ";0", _
        "0;;zero", "0;;;null", "0.00", "#,##0.00", "0.00E+00" _
    )
    For i = LBound(patterns) To UBound(patterns)
        EmitArguments "TOKEN-" & Format$(i + 1, "00"), CStr(patterns(i))
    Next i
End Sub

Private Sub EmitArguments(ByVal caseId As String, ByVal pattern As String)
    Dim value As Variant
    value = 1000
    EmitCase caseId & "-NUM", pattern, value
    value = -1000
    EmitCase caseId & "-NEG", pattern, value
    value = 0
    EmitCase caseId & "-ZERO", pattern, value
    value = DateSerial(2026, 3, 15) + TimeSerial(12, 34, 56)
    EmitCase caseId & "-DATE", pattern, value
    value = CCur("1234.5")
    EmitCase caseId & "-CURRENCY", pattern, value
    value = "Ab9"
    EmitCase caseId & "-STRING", pattern, value
End Sub

Private Sub EmitCase(ByVal caseId As String, ByVal pattern As String, ByVal value As Variant)
    Dim output As Variant, errNo As Long, line As String
    On Error Resume Next
    Err.Clear
    output = Format(value, pattern)
    errNo = Err.Number
    If errNo <> 0 Then
        line = "CASE=" & caseId & " PATTERN=" & Encode(pattern) & _
            " ARGTYPE=" & TypeName(value) & " ERR=" & CStr(errNo)
    Else
        line = "CASE=" & caseId & " PATTERN=" & Encode(pattern) & _
            " ARGTYPE=" & TypeName(value) & " VALUE=" & Encode(CStr(output))
    End If
    Print #resultFile, line
    Debug.Print line
    Err.Clear
    On Error GoTo 0
End Sub

Private Function Encode(ByVal value As String) As String
    Encode = Replace(Replace(Replace(value, vbCr, "<CR>"), vbLf, "<LF>"), vbTab, "<TAB>")
End Function
