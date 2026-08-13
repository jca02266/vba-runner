Attribute VB_Name = "FormatMatrixVerification"
Option Explicit

Private resultFile As Integer
Private Const QUEUE_SOURCE_SHA256 As String = ""

Public Sub RunFormatMatrixVerification()
    Dim path As String
    path = ThisWorkbook.Path & Application.PathSeparator & "FormatMatrix.result"
    resultFile = FreeFile
    Open path For Output As #resultFile
    Print #resultFile, "FORMAT_MATRIX_BEGIN"
    EmitSingleCharacterMatrix
    EmitCompoundTokenMatrix
    EmitTokenizerCornerMatrix
    EmitMixedNullBoundaryMatrix
    EmitScalingBoundaryMatrix
    EmitDateContextBoundaryMatrix
    EmitSelectedSectionBoundaryMatrix
    EmitNamedDateBoundaryMatrix
    EmitRoundingBoundaryMatrix
    EmitNamedBooleanBoundaryMatrix
    EmitScientificLiteralBoundaryMatrix
    EmitEmptyStringSectionBoundaryMatrix
    Print #resultFile, "QUEUE_SOURCE_SHA256=" & QUEUE_SOURCE_SHA256
    Print #resultFile, "FORMAT_MATRIX_COMPLETE=True"
    Close #resultFile
End Sub

Private Sub EmitEmptyStringSectionBoundaryMatrix()
    Print #resultFile, "XL-199"
    Debug.Print "XL-199"
    EmitCase "XL-199-AT-EMPTY", "@;EMPTY", ""
    EmitCase "XL-199-NUM-EMPTY", "0.00;NEG;ZERO;NULL", ""
    EmitCase "XL-199-DATE-EMPTY", "yyyy-mm-dd;EMPTY", ""
    EmitCase "XL-199-NUM-DATE-EMPTY", "0.00;yyyy-mm-dd;EMPTY", ""
    EmitCase "XL-199-NUM-NEG-ZERO-EMPTY", "0.00;NEG;ZERO;EMPTY", ""
End Sub

Private Sub EmitScientificLiteralBoundaryMatrix()
    Print #resultFile, "XL-198"
    Debug.Print "XL-198"
    EmitCase "XL-198-UNESCAPED", "0.00E+00,,", CCur("1000000")
    EmitCase "XL-198-ESCAPED-ONE", "0.00E+00\,", CCur("1000000")
    EmitCase "XL-198-ESCAPED-TWO", "0.00E+00\,\,", CCur("1000000")
End Sub

Private Sub EmitNamedBooleanBoundaryMatrix()
    Dim value As Variant
    Print #resultFile, "XL-186"
    Debug.Print "XL-186"
    value = Empty
    EmitCase "XL-186-EMPTY-TF", "True/False", value
    EmitCase "XL-186-EMPTY-YN", "Yes/No", value
    EmitCase "XL-186-EMPTY-OO", "On/Off", value
    value = Null
    EmitCase "XL-186-NULL-TF", "True/False", value
    EmitCase "XL-186-NULL-YN", "Yes/No", value
    EmitCase "XL-186-NULL-OO", "On/Off", value
    EmitCase "XL-186-STRING1-TF", "True/False", "1"
    EmitCase "XL-186-STRING0-YN", "Yes/No", "0"
    EmitCase "XL-186-STRING-OO", "On/Off", "abc"
End Sub

Private Sub EmitRoundingBoundaryMatrix()
    Print #resultFile, "XL-185"
    Debug.Print "XL-185"
    EmitCase "XL-185-POS-BELOW", "0", 1.499999999
    EmitCase "XL-185-NEG-BELOW", "0", -1.499999999
    EmitCase "XL-185-POS-HALF", "0", 1.5
    EmitCase "XL-185-NEG-HALF", "0", -1.5
    EmitCase "XL-185-POS-DECIMAL", "0.0", 2.349999999
    EmitCase "XL-185-NEG-DECIMAL", "0.0", -2.349999999
End Sub

Private Sub EmitNamedDateBoundaryMatrix()
    Print #resultFile, "XL-184"
    Debug.Print "XL-184"
    EmitCase "XL-184-DATE-GENERALNUMBER", "General Number", DateSerial(2026, 3, 15) + TimeSerial(13, 4, 5)
    EmitCase "XL-184-DATE-CURRENCY", "Currency", DateSerial(2026, 3, 15) + TimeSerial(13, 4, 5)
    EmitCase "XL-184-DATE-FIXED", "Fixed", DateSerial(2026, 3, 15) + TimeSerial(13, 4, 5)
    EmitCase "XL-184-DATE-SCIENTIFIC", "Scientific", DateSerial(2026, 3, 15) + TimeSerial(13, 4, 5)
    EmitCase "XL-184-CURRENCY-GENERALDATE", "General Date", CCur("46096.5")
    EmitCase "XL-184-DECIMAL-GENERALDATE", "General Date", CDec("46096.5")
    EmitCase "XL-184-STRING-GENERALDATE", "General Date", "abc"
End Sub

Private Sub EmitSelectedSectionBoundaryMatrix()
    Print #resultFile, "XL-183"
    Debug.Print "XL-183"
    EmitCase "XL-183-DATESTRING", "0.00;NEG;ZERO;NULL", "2026-03-15"
    EmitCase "XL-183-NUMSTRING", "0.00;NEG;ZERO;NULL", "12"
    EmitCase "XL-183-QUOTED", "0.00;""NEG"";""ZERO"";""NULL""", "2026-03-15"
End Sub

Private Sub EmitDateContextBoundaryMatrix()
    Print #resultFile, "XL-182"
    Debug.Print "XL-182"
    EmitCase "XL-182-HM", "h m", DateSerial(2026, 3, 15) + TimeSerial(12, 34, 56)
    EmitCase "XL-182-HYPHEN", "h-m", DateSerial(2026, 3, 15) + TimeSerial(12, 34, 56)
    EmitCase "XL-182-LITERAL", "hfoo m", DateSerial(2026, 3, 15) + TimeSerial(12, 34, 56)
    EmitCase "XL-182-C", "c", DateSerial(2026, 3, 15) + TimeSerial(12, 34, 56)
End Sub

Private Sub EmitScalingBoundaryMatrix()
    Print #resultFile, "XL-181"
    Debug.Print "XL-181"
    EmitCase "XL-181-POS", "##0,", 1000000
    EmitCase "XL-181-NEG", "##0,", -1234567.89
    EmitCase "XL-181-ZERO", "##0,", 0
    EmitCase "XL-181-DOUBLE", "##0,,", 100000000
    EmitCase "XL-181-GROUP", "#,##0,", 1234567
    EmitCase "XL-181-DECIMAL", "0.00,,", 1234567.89
    EmitCase "XL-181-PERCENT", "0.00,,%", 0.125
    EmitCase "XL-181-PERCENT-SUFFIX", "0.00%,,", 0.125
    EmitCase "XL-181-SCIENTIFIC", "0.00E+00,,", 1000000
    EmitCase "XL-181-SCIENTIFIC-PREFIX", "0.00,,E+00", 1000000
End Sub

Private Sub EmitMixedNullBoundaryMatrix()
    Print #resultFile, "XL-180"
    Debug.Print "XL-180"
    EmitCase "XL-180-POS", "0.00;NEG;ZERO;NULL", "12"
    EmitCase "XL-180-NEG", "0.00;NEG;ZERO;NULL", "-12"
    EmitCase "XL-180-ZERO", "0.00;NEG;ZERO;NULL", "0"
    EmitCase "XL-180-NULL", "0.00;(0.00);-;NULL", Null
    EmitCase "XL-180-EMPTY", "0.00;(0.00);-;EMPTY", Empty
    EmitCase "XL-180-MIX-NUM", "0 @ & !", 123
    EmitCase "XL-180-MIX-STR", "0 @ & !", "AB"
    EmitCase "XL-180-MIX-NONNUM", "0.00 @", "abc"
End Sub

Private Sub EmitTokenizerCornerMatrix()
    Dim patterns As Variant, probeIds As Variant, i As Long
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
    probeIds = Array( _
        "XL-146", "XL-147", "XL-148", "XL-149", "XL-150", "XL-151", "XL-152", "XL-153", _
        "XL-154", "XL-155", "XL-156", "XL-157", "XL-158", "XL-159", "XL-160", "XL-161", _
        "XL-162", "XL-163", "XL-164", "XL-165", "XL-166", "XL-167", "XL-168", "XL-169", _
        "XL-170", "XL-171", "XL-172", "XL-173", "XL-174" _
    )
    For i = LBound(patterns) To UBound(patterns)
        EmitArguments CStr(probeIds(i)), CStr(patterns(i))
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
