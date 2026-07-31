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

Private Type VariableStringRecord
    Id As Long
    Name As String
End Type

Private Type NestedChildRecord
    Id As Long
    Name As String
End Type

Private Type NestedParentRecord
    Header As NestedChildRecord
    Values(0 To 1) As Long
    Tail As String
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
    VerifyVariableStringRecord root & Application.PathSeparator & "XL-008.bin", False
    VerifyVariableStringRecord root & Application.PathSeparator & "XL-008-random.bin", True
    VerifyTextStream root & Application.PathSeparator & "XL-009-unicode.txt", True
    VerifyTextStream root & Application.PathSeparator & "XL-009-ansi.txt", False
    VerifyNestedRecord root & Application.PathSeparator & "XL-010.bin"
    VerifySelectCaseString
    VerifyDecimalIntegerOps
    VerifyCurrencyDivision
    VerifyLongLongStringOps
    VerifyCurrencyStringInputs
    VerifyAppendWidthInitialColumn root & Application.PathSeparator & "XL-018-append-width.bin"
    VerifyCrossHandleLock root & Application.PathSeparator & "XL-019-lock.bin"
    Debug.Print "XL-020 ALREADY_VERIFIED=SECONDERR=70"
    VerifyCDecBoundaries
    VerifyFormatRounding
    VerifySeekBoundaries root & Application.PathSeparator & "XL-024-seek.dat"
    VerifyBinaryTextEof root & Application.PathSeparator & "XL-025-text-eof.dat"
    Debug.Print "XL-023 SKIPPED=逐次モードLock境界はExcelで待機する可能性があるため単発実行"
    Debug.Print "RESULT_ROOT=" & root
End Sub

Public Sub RunExcelSequentialLockVerification()
    Dim root As String
    root = Environ$("TEMP") & Application.PathSeparator & "vba-runner-xl-queue"
    On Error Resume Next
    MkDir root
    On Error GoTo 0
    VerifySequentialLockBoundaries root & Application.PathSeparator & "XL-023-lock-range.dat"
    Debug.Print "XL-023 RESULT_ROOT=" & root
End Sub

Private Sub VerifySelectCaseString()
    Dim value As Variant, result As String
    value = "12"
    Select Case value
        Case 12
            result = "numeric-case"
        Case Else
            result = "else"
    End Select
    Debug.Print "XL-013 SELECT=" & result & " TYPENAME=" & TypeName(value)
End Sub

Private Sub VerifyDecimalIntegerOps()
    Dim value As Variant, quotient As Variant, remainder As Variant, errNo As Long
    value = CDec("79228162514264337593543950335")
    On Error Resume Next
    quotient = value \ CDec("3")
    errNo = Err.Number
    Debug.Print "XL-014 DIVERR=" & CStr(errNo)
    Err.Clear
    remainder = value Mod CDec("100000000000000000")
    errNo = Err.Number
    Debug.Print "XL-014 MODERR=" & CStr(errNo)
    Err.Clear
    Debug.Print "XL-014 NEG=" & CStr(CDec("-5.5") \ CDec("2.5"))
    Debug.Print "XL-014 NEGERR=" & CStr(Err.Number)
    On Error GoTo 0
End Sub

Private Sub VerifyCurrencyDivision()
    Dim value As Currency, divisor As Currency, quotient As Variant, remainder As Variant, errNo As Long
    value = CCur("922337203685477.5807")
    divisor = CCur("7")
    On Error Resume Next
    quotient = value / divisor
    errNo = Err.Number
    Debug.Print "XL-015 DIV=" & CStr(quotient) & " TYPE=" & TypeName(quotient) & " ERR=" & CStr(errNo)
    Err.Clear
    remainder = value Mod divisor
    errNo = Err.Number
    If errNo = 0 Then
        Debug.Print "XL-015 MOD=" & CStr(remainder) & " MODTYPE=" & TypeName(remainder) & " ERR=0"
    Else
        Debug.Print "XL-015 MODERR=" & CStr(errNo)
    End If
    Err.Clear
    On Error GoTo 0
End Sub

Private Sub VerifyLongLongStringOps()
    Dim value As LongLong
    value = CLngLng("9007199254740993")
    Debug.Print "XL-016 ADD=" & CStr(value + "1") & " SUB=" & CStr(value - "1") & _
        " NEG=" & CStr(-value)
End Sub

Private Sub VerifyCurrencyStringInputs()
    Debug.Print "XL-017 EXP=" & CStr(CCur("9.007199254740993125E+14"))
    Debug.Print "XL-017 GROUP=" & CStr(CCur("900,719,925,474,099.3125"))
    Debug.Print "XL-017 PLUS=" & CStr(CCur("+900719925474099.3125"))
End Sub

Private Sub VerifyAppendWidthInitialColumn(ByVal path As String)
    Dim f As Integer
    On Error Resume Next
    Kill path
    On Error GoTo 0
    f = FreeFile
    Open path For Output As #f
    Print #f, "1234";
    Close #f
    f = FreeFile
    Open path For Append As #f
    Width #f, 5
    Print #f, "56"
    Close #f
    PrintBytes "XL-018", path
End Sub

Private Sub VerifyCrossHandleLock(ByVal path As String)
    Dim first As Integer, second As Integer, errNo As Long
    On Error Resume Next
    Kill path
    On Error GoTo 0
    first = FreeFile
    Open path For Random Access Read Write Lock Read Write As #first Len = 4
    second = FreeFile
    On Error Resume Next
    Open path For Random Access Read Write Shared As #second Len = 4
    errNo = Err.Number
    Debug.Print "XL-019 OPEN2ERR=" & CStr(errNo)
    If errNo <> 0 Then
        Err.Clear
        Close #first
        On Error GoTo 0
        Exit Sub
    End If
    Err.Clear
    Lock #first, 1 To 2
    errNo = Err.Number
    Debug.Print "XL-019 FIRSTERR=" & CStr(errNo)
    Err.Clear
    Lock #second, 1 To 2
    Debug.Print "XL-019 SECONDERR=" & CStr(Err.Number)
    Err.Clear
    Unlock #first, 1 To 2
    Close #second
    Close #first
    On Error GoTo 0
End Sub

Private Sub VerifySharedLockRange(ByVal path As String)
    Dim first As Integer, second As Integer, errNo As Long
    On Error Resume Next
    Kill path
    On Error GoTo 0
    first = FreeFile
    Open path For Random Access Read Write Shared As #first Len = 4
    second = FreeFile
    On Error Resume Next
    Open path For Random Access Read Write Shared As #second Len = 4
    errNo = Err.Number
    Debug.Print "XL-020 OPEN2ERR=" & CStr(errNo)
    If errNo <> 0 Then
        Err.Clear
        Close #first
        On Error GoTo 0
        Exit Sub
    End If
    Err.Clear
    Lock #first, 1 To 2
    Debug.Print "XL-020 FIRSTERR=" & CStr(Err.Number)
    Err.Clear
    Lock #second, 2 To 3
    Debug.Print "XL-020 SECONDERR=" & CStr(Err.Number)
    Err.Clear
    Unlock #first, 1 To 2
    Close #second
    Close #first
    On Error GoTo 0
End Sub

Private Sub VerifyCDecBoundaries()
    PrintCDecResult "PLUS", "+1"
    PrintCDecResult "HEX", "&H10"
    PrintCDecResult "OCT", "&O10"
    PrintCDecResult "GROUP", "1,234"
    PrintCDecResult "EXP", "1.2E3"
End Sub

Private Sub PrintCDecResult(ByVal label As String, ByVal text As String)
    Dim value As Variant, errNo As Long
    On Error Resume Next
    Err.Clear
    value = CDec(text)
    errNo = Err.Number
    If errNo = 0 Then
        Debug.Print "XL-021 " & label & " ERR=0 VALUE=" & CStr(value)
    Else
        Debug.Print "XL-021 " & label & " ERR=" & CStr(errNo)
    End If
    On Error GoTo 0
End Sub

Private Sub VerifyFormatRounding()
    Debug.Print "XL-022 CURRENCY=[" & Format(2.675, "Currency") & "]"
    Debug.Print "XL-022 FIXED=[" & Format(2.675, "Fixed") & "]"
    Debug.Print "XL-022 STANDARD=[" & Format(2.675, "Standard") & "]"
    Debug.Print "XL-022 SCIENTIFIC=[" & Format(2.675, "Scientific") & "]"
    Debug.Print "XL-022 SCI09995=[" & Format(0.09995, "Scientific") & "]"
End Sub

Private Sub VerifySequentialLockBoundaries(ByVal path As String)
    Dim mode As Variant, f As Integer
    On Error Resume Next
    Kill path
    On Error GoTo 0
    f = FreeFile
    Open path For Output As #f
    Close #f
    For Each mode In Array("Input", "Output", "Append")
        ProbeSequentialLock path, CStr(mode)
    Next mode
End Sub

Private Sub ProbeSequentialLock(ByVal path As String, ByVal mode As String)
    Dim f As Integer, zeroErr As Long, negativeErr As Long, fractionErr As Long, reverseErr As Long
    If mode = "Input" Then
        f = FreeFile: Open path For Input As #f
    ElseIf mode = "Output" Then
        f = FreeFile: Open path For Output As #f
    Else
        f = FreeFile: Open path For Append As #f
    End If
    On Error Resume Next
    Err.Clear: Lock #f, 0 To 1: zeroErr = Err.Number
    Err.Clear: Lock #f, -1 To 1: negativeErr = Err.Number
    Err.Clear: Lock #f, 1.5 To 2: fractionErr = Err.Number
    Err.Clear: Lock #f, 2 To 1: reverseErr = Err.Number
    Close #f
    On Error GoTo 0
    Debug.Print "XL-023 MODE=" & mode & " ZERO=" & CStr(zeroErr) & _
        " NEG=" & CStr(negativeErr) & " FRACTION=" & CStr(fractionErr) & _
        " REVERSE=" & CStr(reverseErr)
End Sub

Private Sub VerifySeekBoundaries(ByVal path As String)
    Dim mode As Variant
    For Each mode In Array("Input", "Output", "Append", "Random", "Binary")
        ProbeSeek path, CStr(mode)
    Next mode
End Sub

Private Sub ProbeSeek(ByVal path As String, ByVal mode As String)
    Dim f As Integer, seed As Integer, zeroErr As Long, negativeErr As Long, fractionErr As Long, highErr As Long
    On Error Resume Next
    Kill path
    On Error GoTo 0
    If mode = "Input" Then
        seed = FreeFile
        Open path For Output As #seed
        Close #seed
        f = FreeFile: Open path For Input As #f
    ElseIf mode = "Output" Then
        f = FreeFile: Open path For Output As #f
    ElseIf mode = "Append" Then
        f = FreeFile: Open path For Append As #f
    ElseIf mode = "Random" Then
        f = FreeFile: Open path For Random As #f Len = 4
    Else
        f = FreeFile: Open path For Binary As #f
    End If
    On Error Resume Next
    Err.Clear: Seek #f, 0: zeroErr = Err.Number
    Err.Clear: Seek #f, -1: negativeErr = Err.Number
    Err.Clear: Seek #f, 1.5: fractionErr = Err.Number
    Err.Clear: Seek #f, 2147483647: highErr = Err.Number
    Close #f
    On Error GoTo 0
    Debug.Print "XL-024 MODE=" & mode & " ZERO=" & CStr(zeroErr) & _
        " NEG=" & CStr(negativeErr) & " FRACTION=" & CStr(fractionErr) & _
        " HIGH=" & CStr(highErr)
End Sub

Private Sub VerifyBinaryTextEof(ByVal path As String)
    Dim f As Integer, text As String, errNo As Long
    On Error Resume Next
    Kill path
    On Error GoTo 0
    f = FreeFile
    Open path For Output As #f
    Print #f, "alpha"
    Close #f

    f = FreeFile: Open path For Input As #f
    Line Input #f, text
    Debug.Print "XL-025 INPUT LINEEOF=" & CStr(EOF(f))
    On Error Resume Next
    Err.Clear: Line Input #f, text: errNo = Err.Number
    Debug.Print "XL-025 INPUT LINEERR=" & CStr(errNo) & " EOF=" & CStr(EOF(f))
    Close #f
    On Error GoTo 0

    f = FreeFile: Open path For Binary As #f
    Line Input #f, text
    Debug.Print "XL-025 BINARY LINEEOF=" & CStr(EOF(f))
    On Error Resume Next
    Err.Clear: Line Input #f, text: errNo = Err.Number
    Debug.Print "XL-025 BINARY LINEERR=" & CStr(errNo) & " EOF=" & CStr(EOF(f))
    Close #f
    On Error GoTo 0
End Sub

Private Sub VerifyTextStream(ByVal path As String, ByVal unicodeMode As Boolean)
    Dim fso As Object, stream As Object, id As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set stream = fso.CreateTextFile(path, True, unicodeMode)
    stream.Write "Aあ"
    stream.WriteLine "B"
    stream.Close
    id = IIf(unicodeMode, "XL-009-UNICODE", "XL-009-ANSI")
    PrintBytes id, path
End Sub

Private Sub VerifyNestedRecord(ByVal path As String)
    Dim record As NestedParentRecord, f As Integer
    record.Header.Id = &H1020304
    record.Header.Name = "Aあ"
    record.Values(0) = 1
    record.Values(1) = 2
    record.Tail = "Z"
    f = FreeFile
    Open path For Binary Access Write As #f
    Put #f, , record
    Close #f
    PrintBytes "XL-010", path
End Sub

Private Sub VerifyVariableStringRecord(ByVal path As String, ByVal randomMode As Boolean)
    Dim record As VariableStringRecord, f As Integer
    record.Id = &H1020304
    record.Name = "Aあ"
    f = FreeFile
    If randomMode Then
        Open path For Random Access Write As #f Len = 16
        Put #f, 1, record
    Else
        Open path For Binary Access Write As #f
        Put #f, , record
    End If
    Close #f
    PrintBytes IIf(randomMode, "XL-008-RANDOM", "XL-008"), path
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
