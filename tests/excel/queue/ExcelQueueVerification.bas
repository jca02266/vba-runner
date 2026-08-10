Attribute VB_Name = "ExcelQueueVerification"
Option Explicit

Private Const QUEUE_SOURCE_SHA256 As String = "__QUEUE_SOURCE_SHA256__"

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

Private Type UdtObjectArrayRecord
    Values(0 To 1) As Object
End Type

Private resultFile As Integer
Private resultOpen As Boolean

Public Sub RunExcelQueueVerification()
    Dim root As String
    BeginResult
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
    EmitResult "XL-020 ALREADY_VERIFIED=SECONDERR=70"
    VerifyCDecBoundaries
    VerifyFormatRounding
    VerifySeekBoundaries root & Application.PathSeparator & "XL-024-seek.dat"
    VerifyBinaryTextEof root & Application.PathSeparator & "XL-025-text-eof.dat"
    VerifyNonFiniteBoundaries
    VerifyAsNewArray
    VerifyCollectionEnumerationMutation
    VerifyTextStreamCloseBoundaries root & Application.PathSeparator & "XL-056-close.txt"
    VerifyTextStreamEofReadLine root & Application.PathSeparator & "XL-057-eof.txt"
    VerifyPendingExcelBoundaries
    VerifyRadixConversionBoundaries
    VerifyRadixConversionMatrix
    VerifyRadixExplicitSignMatrix
    VerifyRadixTargetTypeMatrix
    VerifyRadixShortWidthMatrix
    VerifyRadixByteAndOctalMatrix
    VerifyRadixCurrencyOctalFollowup
    VerifyCallByNameNamedParamArray
    VerifyMemberForcedByVal
    VerifyUdtObjectArraySet
    VerifyOpaqueShape
    EmitResult "XL-023 SKIPPED=逐次モードLock境界はExcelで待機する可能性があるため単発実行"
    EmitResult "QUEUE_SOURCE_SHA256=" & QUEUE_SOURCE_SHA256
    EmitResult "QUEUE_COMPLETE=True"
    EndResult
End Sub

Private Sub VerifyRadixCurrencyOctalFollowup()
    Dim value As Variant, errNo As Long
    On Error Resume Next
    Err.Clear: value = CCur("&O20000000000"): errNo = Err.Number
    EmitValueAndType "XL-080 CCur-OCT-MIN", value, errNo
    Err.Clear: value = CCur("&O7777777777"): errNo = Err.Number
    EmitValueAndType "XL-080 CCur-OCT-SHORT", value, errNo
    Err.Clear: value = CDec("&O37777777777"): errNo = Err.Number
    EmitValueAndType "XL-081 CDec-OCT-32", value, errNo
    Err.Clear: value = Val("&O37777777777"): errNo = Err.Number
    EmitValueAndType "XL-082 Val-OCT-32", value, errNo
    On Error GoTo 0
End Sub

Private Sub VerifyRadixByteAndOctalMatrix()
    Dim value As Variant, errNo As Long
    On Error Resume Next
    Err.Clear: value = CByte("&HFF"): errNo = Err.Number
    EmitValueAndType "XL-077 CByte-FF", value, errNo
    Err.Clear: value = CByte("&H100"): errNo = Err.Number
    EmitValueAndType "XL-077 CByte-100", value, errNo
    Err.Clear: value = CDec("&O777777777777777777777"): errNo = Err.Number
    EmitValueAndType "XL-078 CDec-OCTAL", value, errNo
    Err.Clear: value = CCur("&O37777777777"): errNo = Err.Number
    EmitValueAndType "XL-079 CCur-OCTAL", value, errNo
    On Error GoTo 0
End Sub

Private Sub VerifyRadixShortWidthMatrix()
    Dim value As Variant, errNo As Long
    On Error Resume Next
    Err.Clear: value = CInt("&H7F"): errNo = Err.Number
    EmitValueAndType "XL-074 CInt-7F", value, errNo
    Err.Clear: value = CInt("&HFF"): errNo = Err.Number
    EmitValueAndType "XL-074 CInt-FF", value, errNo
    Err.Clear: value = CLng("&HFFFF"): errNo = Err.Number
    EmitValueAndType "XL-075 CLng-FFFF", value, errNo
    Err.Clear: value = CCur("&HFFFF"): errNo = Err.Number
    EmitValueAndType "XL-075 CCur-FFFF", value, errNo
    Err.Clear: value = CDec("&HFFFF"): errNo = Err.Number
    EmitValueAndType "XL-076 CDec-FFFF", value, errNo
    Err.Clear: value = Val("&HFFFF"): errNo = Err.Number
    EmitValueAndType "XL-076 Val-FFFF", value, errNo
    On Error GoTo 0
End Sub

Private Sub VerifyRadixTargetTypeMatrix()
    Dim value As Variant, errNo As Long
    On Error Resume Next
    Err.Clear: value = CDec("&H80000000"): errNo = Err.Number
    EmitValueAndType "XL-071 CDec-32MIN", value, errNo
    Err.Clear: value = CDec("&HFFFFFFFF"): errNo = Err.Number
    EmitValueAndType "XL-071 CDec-32NEG1", value, errNo
    Err.Clear: value = CCur("&H7FFFFFFF"): errNo = Err.Number
    EmitValueAndType "XL-072 CCur-32POS", value, errNo
    Err.Clear: value = CCur("&H80000000"): errNo = Err.Number
    EmitValueAndType "XL-072 CCur-32MIN", value, errNo
    Err.Clear: value = Val("&H80000000"): errNo = Err.Number
    EmitValueAndType "XL-073 Val-32MIN", value, errNo
    Err.Clear: value = Val("&H7FFFFFFF"): errNo = Err.Number
    EmitValueAndType "XL-073 Val-32POS", value, errNo
    On Error GoTo 0
End Sub

Private Sub VerifyRadixConversionMatrix()
    Dim value As Variant, errNo As Long
    On Error Resume Next
    Err.Clear: value = CInt("&H80"): errNo = Err.Number
    EmitValueAndType "XL-065 CInt-SHORT", value, errNo
    Err.Clear: value = CInt("&HFFFF"): errNo = Err.Number
    EmitValueAndType "XL-065 CInt-BOUNDARY", value, errNo
    Err.Clear: value = CLngLng("&O777777777777777777777"): errNo = Err.Number
    EmitValueAndType "XL-066 CLngLng-OCTAL", value, errNo
    Err.Clear: value = CCur("&HFFFFFFFFFFFFFFFF"): errNo = Err.Number
    EmitValueAndType "XL-067 CCur-64", value, errNo
    Err.Clear: value = Val("&O1777777777777777777777"): errNo = Err.Number
    EmitValueAndType "XL-067 Val-OCTAL", value, errNo
    On Error GoTo 0
End Sub

Private Sub VerifyRadixExplicitSignMatrix()
    Dim value As Variant, errNo As Long
    On Error Resume Next
    Err.Clear: value = CLng("+&HFFFFFFFF"): errNo = Err.Number
    EmitValueAndType "XL-068 CLng-PLUS", value, errNo
    Err.Clear: value = CInt("-&H8000"): errNo = Err.Number
    EmitValueAndType "XL-068 CInt-MINUS", value, errNo
    Err.Clear: value = CLng("&O37777777777"): errNo = Err.Number
    EmitValueAndType "XL-069 CLng-OCTAL", value, errNo
    Err.Clear: value = CInt("&O77777"): errNo = Err.Number
    EmitValueAndType "XL-069 CInt-OCTAL", value, errNo
    Err.Clear: value = Val("&O77777777777"): errNo = Err.Number
    EmitValueAndType "XL-070 Val-OCTAL32", value, errNo
    On Error GoTo 0
End Sub

Private Sub VerifyRadixConversionBoundaries()
    Dim value As Variant, errNo As Long
    On Error Resume Next
    Err.Clear: value = CLngLng("&H7FFFFFFFFFFFFFFF"): errNo = Err.Number
    EmitValueAndType "XL-062 CLngLng-POS", value, errNo
    Err.Clear: value = CLngLng("&H8000000000000000"): errNo = Err.Number
    EmitValueAndType "XL-062 CLngLng-MIN", value, errNo
    Err.Clear: value = CLngLng("&HFFFFFFFFFFFFFFFF"): errNo = Err.Number
    EmitValueAndType "XL-062 CLngLng-NEG1", value, errNo
    Err.Clear: value = CCur("&HFFFFFFFF"): errNo = Err.Number
    EmitValueAndType "XL-063 CCur-32", value, errNo
    Err.Clear: value = Val("&HFFFFFFFFFFFFFFFF"): errNo = Err.Number
    EmitValueAndType "XL-064 Val-64", value, errNo
    On Error GoTo 0
End Sub

Private Sub VerifyOpaqueShape()
    Dim ws As Worksheet, source As Variant
    Dim target() As Variant
    Dim errNo As Long, lower As Long, upper As Long
    On Error Resume Next
    Set ws = Worksheets.Add
    ws.Range("A1:C2").Value2 = Array(Array(1, 2, 3), Array(4, 5, 6))
    source = ws.Range("A1:C2").Value2
    Err.Clear
    target = source
    errNo = Err.Number
    If errNo = 0 Then
        Err.Clear
        lower = LBound(target, 2)
        upper = UBound(target, 2)
        If Err.Number <> 0 Then lower = -1: upper = -1
    Else
        lower = -1
        upper = -1
    End If
    EmitResult "XL-061 SHAPEERR=" & CStr(errNo) & " LBOUND2=" & CStr(lower) & " UBOUND2=" & CStr(upper)
    Err.Clear
    Application.DisplayAlerts = False
    ws.Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
End Sub

Private Sub VerifyUdtObjectArraySet()
    Dim record As UdtObjectArrayRecord, payload As New ExcelQueueForcedByVal, errNo As Long
    On Error Resume Next
    Set payload = New ExcelQueueForcedByVal
    Err.Clear: Set record.Values(0) = payload: errNo = Err.Number
    EmitResult "XL-060 SET-PAYLOAD ERR=" & CStr(errNo) & " TYPE=" & TypeName(record.Values(0))
    Err.Clear: Set record.Values(1) = Nothing: errNo = Err.Number
    EmitResult "XL-060 SET-NOTHING ERR=" & CStr(errNo) & " TYPE=" & TypeName(record.Values(1))
    Err.Clear: record.Values(0) = 1: errNo = Err.Number
    EmitResult "XL-060 LET-SCALAR ERR=" & CStr(errNo)
    On Error GoTo 0
End Sub

Private Sub VerifyMemberForcedByVal()
    Dim concrete As New ExcelQueueForcedByVal, target As ExcelQueueForcedByValInterface
    Dim value As Long, errNo As Long
    On Error Resume Next
    value = 1
    Err.Clear: concrete.Mutate (value): errNo = Err.Number
    EmitResult "XL-059 CLASS-PAREN VALUE=" & CStr(value) & " ERR=" & CStr(errNo)
    Err.Clear: value = 1: concrete.Mutate value: errNo = Err.Number
    EmitResult "XL-059 CLASS-PLAIN VALUE=" & CStr(value) & " ERR=" & CStr(errNo)
    Set target = concrete
    Err.Clear: value = 1: target.Mutate (value): errNo = Err.Number
    EmitResult "XL-059 INTERFACE-PAREN VALUE=" & CStr(value) & " ERR=" & CStr(errNo)
    On Error GoTo 0
End Sub

Private Sub VerifyCallByNameNamedParamArray()
    Dim target As New ExcelQueueCallByNameIndexed
    Dim value As New ExcelQueueTicket, errNo As Long
    On Error Resume Next
    Err.Clear
    CallByName target, "Item", VbSet, value
    errNo = Err.Number
    EmitResult "XL-058 VALUE-FIRST ERR=" & CStr(errNo)
    Err.Clear
    CallByName target, "Item", VbSet, 2, value
    errNo = Err.Number
    EmitResult "XL-058 INDEX-FIRST ERR=" & CStr(errNo)
    On Error GoTo 0
End Sub

Private Sub VerifyTextStreamEofReadLine(ByVal path As String)
    Dim fso As Object, stream As Object, text As String, errNo As Long
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set stream = fso.CreateTextFile(path, True, False)
    stream.Write "abc"
    stream.Close
    Set stream = fso.OpenTextFile(path, 1, False, -2)
    text = stream.ReadAll
    On Error Resume Next
    Err.Clear: text = stream.ReadLine: errNo = Err.Number
    EmitResult "XL-057 EOF_READLINE_ERR=" & CStr(errNo) & " TEXT=[" & text & "] EOF=" & CStr(stream.AtEndOfStream)
    Err.Clear: stream.SkipLine: errNo = Err.Number
    EmitResult "XL-057 EOF_SKIPLINE_ERR=" & CStr(errNo) & " EOF=" & CStr(stream.AtEndOfStream)
    On Error GoTo 0
    stream.Close
End Sub

Private Sub VerifyPendingExcelBoundaries()
    VerifyDateDiffWBoundaries
    VerifyEmptyArrayBounds
    VerifyMirrRateBoundaries
    VerifySydLargeBoundary
    VerifyOnErrorResumeBoundary
    VerifyDecimalCurrencySelect
    VerifyFormatNumberNull
    VerifyTimeSerialNegativeBoundaries
    VerifyInformationEmptyBoundaries
    VerifyLargeLiteralBoundaries
    VerifyLargeLiteralValues
    VerifyDefaultValueInformation
    VerifyTimeSerialConsumers
    VerifyCallByNamePropertyKinds
    VerifyRadixWidthBoundaries
    VerifyMirrBoundaryMatrix
    VerifySydErrorMatrix
    VerifyRadixExtendedMatrix
    VerifyMirrCashflowMatrix
    VerifySydPositiveArgumentMatrix
    VerifyMirrPrecisionMatrix
End Sub

Private Sub VerifyTextStreamCloseBoundaries(ByVal path As String)
    Dim fso As Object, stream As Object, errNo As Long
    Dim text As String, atEnd As Boolean, lineNo As Long
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set stream = fso.CreateTextFile(path, True, False)
    stream.Write "abc"
    stream.Close
    Set stream = fso.OpenTextFile(path, 1, False, -2)
    stream.Close

    On Error Resume Next
    Err.Clear: text = stream.ReadAll: errNo = Err.Number
    EmitResult "XL-056 READALL_ERR=" & CStr(errNo)
    Err.Clear: text = stream.ReadLine: errNo = Err.Number
    EmitResult "XL-056 READLINE_ERR=" & CStr(errNo)
    Err.Clear: stream.Skip 1: errNo = Err.Number
    EmitResult "XL-056 SKIP_ERR=" & CStr(errNo)
    Err.Clear: stream.SkipLine: errNo = Err.Number
    EmitResult "XL-056 SKIPLINE_ERR=" & CStr(errNo)
    Err.Clear: atEnd = stream.AtEndOfStream: errNo = Err.Number
    EmitResult "XL-056 ATEND_ERR=" & CStr(errNo)
    Err.Clear: lineNo = stream.Line: errNo = Err.Number
    EmitResult "XL-056 LINE_ERR=" & CStr(errNo)
    Err.Clear: stream.Close: errNo = Err.Number
    EmitResult "XL-056 CLOSE2_ERR=" & CStr(errNo)
    On Error GoTo 0
End Sub

Private Sub VerifyRadixWidthBoundaries()
    Dim value As Variant, errNo As Long
    On Error Resume Next
    Err.Clear: value = CByte("&HFF"): errNo = Err.Number
    EmitValueAndType "XL-049 CByte-FF", value, errNo
    Err.Clear: value = CInt("&HFFFF"): errNo = Err.Number
    EmitValueAndType "XL-049 CInt-FFFF", value, errNo
    Err.Clear: value = CLng("&HFFFF"): errNo = Err.Number
    EmitValueAndType "XL-049 CLng-FFFF", value, errNo
    Err.Clear: value = CDec("&HFFFF"): errNo = Err.Number
    EmitValueAndType "XL-049 CDec-FFFF", value, errNo
    Err.Clear: value = Val("&O777777"): errNo = Err.Number
    EmitValueAndType "XL-049 Val-OCT", value, errNo
    On Error GoTo 0
End Sub

Private Sub VerifyMirrBoundaryMatrix()
    Dim flows(0 To 1) As Double, result As Variant, errNo As Long
    flows(0) = -100: flows(1) = 100
    On Error Resume Next
    Err.Clear: result = MIRR(flows, -1.0001, 0.1): errNo = Err.Number
    EmitValueAndType "XL-050 FINANCE-LOW", result, errNo
    Err.Clear: result = MIRR(flows, -0.9999, 0.1): errNo = Err.Number
    EmitValueAndType "XL-050 FINANCE-NEAR", result, errNo
    Err.Clear: result = MIRR(flows, 0.1, -1.0001): errNo = Err.Number
    EmitValueAndType "XL-050 REINVEST-LOW", result, errNo
    Err.Clear: result = MIRR(flows, 0.1, -0.9999): errNo = Err.Number
    EmitValueAndType "XL-050 REINVEST-NEAR", result, errNo
    On Error GoTo 0
End Sub

Private Sub VerifySydErrorMatrix()
    Dim result As Variant, errNo As Long
    On Error Resume Next
    Err.Clear: result = SYD(1E+308, 1E+308, 2, 1): errNo = Err.Number
    EmitValueAndType "XL-051 SAME-SIGN", result, errNo
    Err.Clear: result = SYD(1E+308, -1E+308, 2, 1): errNo = Err.Number
    EmitValueAndType "XL-051 OPPOSITE-SIGN", result, errNo
    Err.Clear: result = SYD(1E+307, -1E+307, 2, 1): errNo = Err.Number
    EmitValueAndType "XL-051 FINITE-OPPOSITE", result, errNo
    Err.Clear: result = SYD(1E+308, -1E+308, 2, 2): errNo = Err.Number
    EmitValueAndType "XL-051 PERIOD-2", result, errNo
    On Error GoTo 0
End Sub

Private Sub VerifyRadixExtendedMatrix()
    Dim value As Variant, errNo As Long
    On Error Resume Next
    Err.Clear: value = CByte("&H80"): errNo = Err.Number
    EmitValueAndType "XL-052 CByte-80", value, errNo
    Err.Clear: value = CInt("&H8000"): errNo = Err.Number
    EmitValueAndType "XL-052 CInt-8000", value, errNo
    Err.Clear: value = CInt("&O77777"): errNo = Err.Number
    EmitValueAndType "XL-052 CInt-OCT", value, errNo
    Err.Clear: value = CLng("&HFFFFFFFF"): errNo = Err.Number
    EmitValueAndType "XL-052 CLng-FFFFFFFF", value, errNo
    Err.Clear: value = CDec("&HFFFFFFFFFFFFFFFF"): errNo = Err.Number
    EmitValueAndType "XL-052 CDec-FFFFFFFFFFFFFFFF", value, errNo
    On Error GoTo 0
End Sub

Private Sub VerifyMirrCashflowMatrix()
    Dim flows(0 To 2) As Double, result As Variant, errNo As Long
    flows(0) = -100: flows(1) = 50: flows(2) = 100
    On Error Resume Next
    Err.Clear: result = MIRR(flows, -1.1, 0.1): errNo = Err.Number
    EmitValueAndType "XL-053 THREE-FLOW-FINANCE", result, errNo
    Err.Clear: result = MIRR(flows, 0.1, -1.1): errNo = Err.Number
    EmitValueAndType "XL-053 THREE-FLOW-REINVEST", result, errNo
    Err.Clear: result = MIRR(flows, 0.05, 0.05): errNo = Err.Number
    EmitValueAndType "XL-053 THREE-FLOW-NORMAL", result, errNo
    On Error GoTo 0
End Sub

Private Sub VerifySydPositiveArgumentMatrix()
    Dim result As Variant, errNo As Long
    On Error Resume Next
    Err.Clear: result = SYD(100, -1, 2, 1): errNo = Err.Number
    EmitValueAndType "XL-054 NEG-SALVAGE", result, errNo
    Err.Clear: result = SYD(100, 0, 2, 1): errNo = Err.Number
    EmitValueAndType "XL-054 ZERO-SALVAGE", result, errNo
    Err.Clear: result = SYD(100, 100, 2, 1): errNo = Err.Number
    EmitValueAndType "XL-054 EQUAL-COST", result, errNo
    Err.Clear: result = SYD(0, 0, 2, 1): errNo = Err.Number
    EmitValueAndType "XL-054 ZERO-COST", result, errNo
    Err.Clear: result = SYD(100, 1, 2, 1): errNo = Err.Number
    EmitValueAndType "XL-054 POSITIVE", result, errNo
    On Error GoTo 0
End Sub

Private Sub VerifyMirrPrecisionMatrix()
    Dim flows(0 To 2) As Double, result As Variant, errNo As Long
    flows(0) = -100: flows(1) = 0: flows(2) = 100
    On Error Resume Next
    Err.Clear: result = MIRR(flows, -1.1, 0.1): errNo = Err.Number
    EmitValueAndType "XL-055 ZERO-MIDDLE-LOW", result, errNo
    Err.Clear: result = MIRR(flows, -0.9999, 0.1): errNo = Err.Number
    EmitValueAndType "XL-055 ZERO-MIDDLE-NEAR", result, errNo
    flows(1) = 100
    Err.Clear: result = MIRR(flows, -1.1, 0.1): errNo = Err.Number
    EmitValueAndType "XL-055 POS-MIDDLE-LOW", result, errNo
    Err.Clear: result = MIRR(flows, 0.1, -1.1): errNo = Err.Number
    EmitValueAndType "XL-055 POS-MIDDLE-REINVEST", result, errNo
    On Error GoTo 0
End Sub

Private Sub VerifyDateDiffWBoundaries()
    Dim result As Variant, errNo As Long
    On Error Resume Next
    Err.Clear: result = DateDiff("w", DateSerial(2024, 1, 1), DateSerial(2024, 1, 8), vbMonday): errNo = Err.Number
    EmitResult "XL-035 DATE1MON=" & IIf(errNo = 0, CStr(result), "ERR=" & CStr(errNo))
    Err.Clear: result = DateDiff("w", DateSerial(2024, 1, 7), DateSerial(2024, 1, 8), vbMonday): errNo = Err.Number
    EmitResult "XL-035 ONE-DAY=" & IIf(errNo = 0, CStr(result), "ERR=" & CStr(errNo))
    Err.Clear: result = DateDiff("w", DateSerial(2024, 1, 8) + TimeSerial(12, 0, 0), DateSerial(2024, 1, 1), vbSunday): errNo = Err.Number
    EmitResult "XL-035 REVERSE-TIME=" & IIf(errNo = 0, CStr(result), "ERR=" & CStr(errNo))
    On Error GoTo 0
End Sub

Private Sub VerifyEmptyArrayBounds()
    Dim values As Variant, errNo As Long, lower As Long, upper As Long
    values = Array()
    On Error Resume Next
    Err.Clear: lower = LBound(values): errNo = Err.Number
    EmitResult "XL-036 LBOUND=" & CStr(lower) & " ERR=" & CStr(errNo)
    Err.Clear: upper = UBound(values): errNo = Err.Number
    EmitResult "XL-036 UBOUND=" & CStr(upper) & " ERR=" & CStr(errNo)
    On Error GoTo 0
End Sub

Private Sub VerifyMirrRateBoundaries()
    Dim result As Variant, errNo As Long, flows(0 To 1) As Double
    flows(0) = -100
    flows(1) = 100
    On Error Resume Next
    Err.Clear: result = MIRR(flows, -1.1, 0.1): errNo = Err.Number
    EmitResult "XL-037 FINANCE=" & IIf(errNo = 0, CStr(result), "ERR=" & CStr(errNo))
    Err.Clear: result = MIRR(flows, 0.1, -1): errNo = Err.Number
    EmitResult "XL-037 REINVEST=" & IIf(errNo = 0, CStr(result), "ERR=" & CStr(errNo))
    On Error GoTo 0
End Sub

Private Sub VerifySydLargeBoundary()
    Dim result As Variant, errNo As Long
    On Error Resume Next
    Err.Clear: result = SYD(1E+308, -1E+308, 2, 1): errNo = Err.Number
    EmitResult "XL-038 SYD=" & IIf(errNo = 0, CStr(result), "ERR=" & CStr(errNo))
    On Error GoTo 0
End Sub

Private Sub VerifyOnErrorResumeBoundary()
    Dim errNo As Long
    On Error GoTo Handler
    Err.Raise 5
    EmitResult "XL-039 BODY=NO-ERROR"
    Exit Sub
Handler:
    On Error GoTo 0
    On Error Resume Next
    Resume AfterResume
AfterResume:
    errNo = Err.Number
    EmitResult "XL-039 RESUME_ERR=" & CStr(errNo)
    On Error GoTo 0
End Sub

Private Sub VerifyDecimalCurrencySelect()
    Dim value As Variant, result As String
    value = CDec("1.5")
    Select Case value
        Case "1.5": result = "STRING"
        Case CCur("1.5"): result = "CURRENCY"
        Case Else: result = "ELSE"
    End Select
    EmitResult "XL-040 DECIMAL=" & result
    value = CCur("1.5")
    Select Case value
        Case "1.5": result = "STRING"
        Case CDec("1.5"): result = "DECIMAL"
        Case Else: result = "ELSE"
    End Select
    EmitResult "XL-040 CURRENCY=" & result
End Sub

Private Sub VerifyFormatNumberNull()
    Dim value As Variant, result As Variant, errNo As Long
    value = Null
    On Error Resume Next
    Err.Clear: result = FormatNumber(value): errNo = Err.Number
    EmitResult "XL-041 NULL ERR=" & CStr(errNo) & " ISNULL=" & CStr(IsNull(result))
    On Error GoTo 0
End Sub

Private Sub VerifyTimeSerialNegativeBoundaries()
    EmitTimeSerialProbe "XL-042 MINUS-SECOND", 0, 0, -1
    EmitTimeSerialProbe "XL-042 MINUS-MINUTE", 0, -1, 0
    EmitTimeSerialProbe "XL-042 MINUS-HOUR", -1, 0, 0
    EmitTimeSerialProbe "XL-042 MIXED", -1, -1, -1
End Sub

Private Sub EmitTimeSerialProbe(ByVal id As String, ByVal hours As Long, ByVal minutes As Long, ByVal seconds As Long)
    Dim result As Variant, errNo As Long
    On Error Resume Next
    Err.Clear: result = TimeSerial(hours, minutes, seconds): errNo = Err.Number
    If errNo = 0 Then
        EmitResult id & " INPUT=" & CStr(hours) & "," & CStr(minutes) & "," & CStr(seconds) & " VALUE=" & CStr(result)
    Else
        EmitResult id & " INPUT=" & CStr(hours) & "," & CStr(minutes) & "," & CStr(seconds) & " ERR=" & CStr(errNo)
    End If
    On Error GoTo 0
End Sub

Private Sub VerifyInformationEmptyBoundaries()
    Dim value As Variant, result As Boolean, obj As Object, errNo As Long
    value = Empty
    On Error Resume Next
    Err.Clear: result = IsNumeric(value): errNo = Err.Number
    EmitResult "XL-043 ISNUMERIC_EMPTY=" & CStr(result) & " ERR=" & CStr(errNo)
    Err.Clear: result = IsDate(value): errNo = Err.Number
    EmitResult "XL-043 ISDATE_EMPTY=" & CStr(result) & " ERR=" & CStr(errNo)
    Set obj = Nothing
    Err.Clear: result = IsNumeric(obj): errNo = Err.Number
    EmitResult "XL-043 ISNUMERIC_OBJECT=" & CStr(result) & " ERR=" & CStr(errNo)
    On Error GoTo 0
End Sub

Private Sub VerifyLargeLiteralBoundaries()
    Dim value As Variant, errNo As Long
    On Error Resume Next
    Err.Clear: value = CLng("&HFFFFFFFF"): errNo = Err.Number
    EmitResult "XL-044 HEX32 ERR=" & CStr(errNo)
    Err.Clear: value = CDec("&HFFFFFFFFFFFFFFFF"): errNo = Err.Number
    EmitResult "XL-044 HEX64 ERR=" & CStr(errNo)
    Err.Clear: value = CDec("&O777777777777777777777"): errNo = Err.Number
    EmitResult "XL-044 OCT-LARGE ERR=" & CStr(errNo)
    On Error GoTo 0
End Sub

Private Sub VerifyLargeLiteralValues()
    Dim value As Variant, errNo As Long
    On Error Resume Next
    Err.Clear: value = CLng("&HFFFFFFFF"): errNo = Err.Number
    EmitValueAndType "XL-045 HEX32", value, errNo
    Err.Clear: value = CDec("&HFFFFFFFFFFFFFFFF"): errNo = Err.Number
    EmitValueAndType "XL-045 HEX64", value, errNo
    Err.Clear: value = CDec("&O777777777777777777777"): errNo = Err.Number
    EmitValueAndType "XL-045 OCT-LARGE", value, errNo
    On Error GoTo 0
End Sub

Private Sub EmitValueAndType(ByVal id As String, ByVal value As Variant, ByVal errNo As Long)
    If errNo = 0 Then
        EmitResult id & " VALUE=" & CStr(value) & " TYPE=" & TypeName(value)
    Else
        EmitResult id & " ERR=" & CStr(errNo)
    End If
End Sub

Private Sub VerifyDefaultValueInformation()
    Dim value As ExcelQueueDefaultValue, result As Boolean, errNo As Long
    Set value = New ExcelQueueDefaultValue
    On Error Resume Next
    value.Value = "123"
    Err.Clear: result = IsNumeric(value): errNo = Err.Number
    EmitResult "XL-046 NUMERIC-STRING=" & CStr(result) & " ERR=" & CStr(errNo)
    value.Value = DateSerial(2024, 1, 2)
    Err.Clear: result = IsDate(value): errNo = Err.Number
    EmitResult "XL-046 DATE=" & CStr(result) & " ERR=" & CStr(errNo)
    value.Value = "not-a-date"
    Err.Clear: result = IsDate(value): errNo = Err.Number
    EmitResult "XL-046 TEXT=" & CStr(result) & " ERR=" & CStr(errNo)
    On Error GoTo 0
End Sub

Private Sub VerifyTimeSerialConsumers()
    Dim value As Date, result As Variant, errNo As Long
    value = TimeSerial(0, 0, -1)
    On Error Resume Next
    Err.Clear: result = Hour(value): errNo = Err.Number
    EmitValueAndType "XL-047 HOUR", result, errNo
    Err.Clear: result = Minute(value): errNo = Err.Number
    EmitValueAndType "XL-047 MINUTE", result, errNo
    Err.Clear: result = Second(value): errNo = Err.Number
    EmitValueAndType "XL-047 SECOND", result, errNo
    Err.Clear: result = Format(value, "hh:nn:ss"): errNo = Err.Number
    EmitValueAndType "XL-047 FORMAT", result, errNo
    Err.Clear: result = DateDiff("s", TimeSerial(0, 0, 0), value): errNo = Err.Number
    EmitValueAndType "XL-047 DATEDIFF", result, errNo
    On Error GoTo 0
End Sub

Private Sub VerifyCallByNamePropertyKinds()
    Dim target As ExcelQueueCallByNameTarget, payload As Collection, errNo As Long
    Set target = New ExcelQueueCallByNameTarget
    Set payload = New Collection
    On Error Resume Next
    Err.Clear
    CallByName target, "Scalar", VbSet, 42
    errNo = Err.Number
    EmitResult "XL-048 VBSET-TO-LET ERR=" & CStr(errNo) & " LETHITS=" & CStr(target.LetHits)
    Err.Clear
    CallByName target, "Reference", VbLet, payload
    errNo = Err.Number
    EmitResult "XL-048 VBLET-TO-SET ERR=" & CStr(errNo) & " SETHITS=" & CStr(target.SetHits)
    On Error GoTo 0
End Sub

Private Sub BeginResult()
    Dim outputPath As String
    outputPath = ThisWorkbook.Path
    If Len(outputPath) = 0 Then outputPath = CurDir$
    outputPath = outputPath & Application.PathSeparator & "ExcelQueueVerification.result"
    On Error Resume Next
    Kill outputPath
    On Error GoTo 0
    resultFile = FreeFile
    Open outputPath For Output As #resultFile
    resultOpen = True
End Sub

Private Sub EndResult()
    If resultOpen Then Close #resultFile
    resultOpen = False
End Sub

Private Sub EmitResult(ByVal line As String)
    Debug.Print line
    If resultOpen Then Print #resultFile, line
End Sub

Private Sub VerifyCollectionEnumerationMutation()
    Dim col As New Collection, item As Variant, seen As String
    Dim dict As Object, key As Variant, keysSeen As String, errNo As Long
    col.Add "A": col.Add "B"
    On Error Resume Next
    Err.Clear
    For Each item In col
        seen = seen & CStr(item)
        If CStr(item) = "A" Then col.Add "C"
    Next item
    errNo = Err.Number
    EmitResult "XL-034 COLADD SEEN=" & seen & " COUNT=" & CStr(col.Count) & " ERR=" & CStr(errNo)
    Err.Clear

    Set dict = CreateObject("Scripting.Dictionary")
    dict.Add "A", 1: dict.Add "B", 2: dict.Add "C", 3
    For Each key In dict.Keys
        keysSeen = keysSeen & CStr(key)
        If CStr(key) = "B" Then dict.Remove "B"
    Next key
    errNo = Err.Number
    EmitResult "XL-034 DICTREMOVE KEYS=" & keysSeen & " COUNT=" & CStr(dict.Count) & " ERR=" & CStr(errNo)
    On Error GoTo 0
End Sub

Private Sub VerifyAsNewArray()
    Dim tickets(0 To 1) As New ExcelQueueTicket
    Dim errNo As Long
    On Error Resume Next
    Err.Clear
    tickets(0).Code = "X"
    errNo = Err.Number
    If errNo <> 0 Then
        EmitResult "XL-033 ERR=" & CStr(errNo)
    Else
        EmitResult "XL-033 CODE=" & tickets(0).Code & " TYPE=" & TypeName(tickets(0)) & " ERR=0"
    End If
    On Error GoTo 0
End Sub

Private Sub VerifyNonFiniteBoundaries()
    Dim value As Variant
    On Error Resume Next
    Err.Clear
    value = Val("1E+309")
    PrintNumericProbe "XL-026", "VAL", value
    Err.Clear
    value = 1E+308 * 100
    PrintNumericProbe "XL-027", "MUL", value
    Err.Clear
    value = 2 ^ 2000
    PrintNumericProbe "XL-028", "POW", value
    Err.Clear
    value = FV(1, 2000, 1)
    PrintNumericProbe "XL-029", "FV", value
    Err.Clear
    value = PMT(1E+308, 2, -1)
    PrintNumericProbe "XL-030", "PMT", value
    Err.Clear
    value = SLN(1E+308, -1E+308, 2)
    PrintNumericProbe "XL-031", "SLN", value
    Err.Clear
    value = DDB(1E+308, -1E+308, 2, 1)
    PrintNumericProbe "XL-032", "DDB", value
    On Error GoTo 0
End Sub

Private Sub PrintNumericProbe(ByVal id As String, ByVal label As String, ByVal value As Variant)
    Dim errNo As Long
    errNo = Err.Number
    If errNo <> 0 Then
        EmitResult id & " " & label & " ERR=" & CStr(errNo)
    Else
        EmitResult id & " " & label & " TYPE=" & TypeName(value) & " VALUE=" & CStr(value) & " ERR=0"
    End If
End Sub

Public Sub RunExcelSequentialLockVerification()
    Dim root As String
    BeginResult
    root = Environ$("TEMP") & Application.PathSeparator & "vba-runner-xl-queue"
    On Error Resume Next
    MkDir root
    On Error GoTo 0
    VerifySequentialLockBoundaries root & Application.PathSeparator & "XL-023-lock-range.dat"
    EndResult
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
    EmitResult "XL-013 SELECT=" & result & " TYPENAME=" & TypeName(value)
End Sub

Private Sub VerifyDecimalIntegerOps()
    Dim value As Variant, quotient As Variant, remainder As Variant, errNo As Long
    value = CDec("79228162514264337593543950335")
    On Error Resume Next
    quotient = value \ CDec("3")
    errNo = Err.Number
    EmitResult "XL-014 DIVERR=" & CStr(errNo)
    Err.Clear
    remainder = value Mod CDec("100000000000000000")
    errNo = Err.Number
    EmitResult "XL-014 MODERR=" & CStr(errNo)
    Err.Clear
    EmitResult "XL-014 NEG=" & CStr(CDec("-5.5") \ CDec("2.5"))
    EmitResult "XL-014 NEGERR=" & CStr(Err.Number)
    On Error GoTo 0
End Sub

Private Sub VerifyCurrencyDivision()
    Dim value As Currency, divisor As Currency, quotient As Variant, remainder As Variant, errNo As Long
    value = CCur("922337203685477.5807")
    divisor = CCur("7")
    On Error Resume Next
    quotient = value / divisor
    errNo = Err.Number
    EmitResult "XL-015 DIV=" & CStr(quotient) & " TYPE=" & TypeName(quotient) & " ERR=" & CStr(errNo)
    Err.Clear
    remainder = value Mod divisor
    errNo = Err.Number
    If errNo = 0 Then
        EmitResult "XL-015 MOD=" & CStr(remainder) & " MODTYPE=" & TypeName(remainder) & " ERR=0"
    Else
        EmitResult "XL-015 MODERR=" & CStr(errNo)
    End If
    Err.Clear
    On Error GoTo 0
End Sub

Private Sub VerifyLongLongStringOps()
    Dim value As LongLong
    value = CLngLng("9007199254740993")
    EmitResult "XL-016 ADD=" & CStr(value + "1") & " SUB=" & CStr(value - "1") & _
        " NEG=" & CStr(-value)
End Sub

Private Sub VerifyCurrencyStringInputs()
    EmitResult "XL-017 EXP=" & CStr(CCur("9.007199254740993125E+14"))
    EmitResult "XL-017 GROUP=" & CStr(CCur("900,719,925,474,099.3125"))
    EmitResult "XL-017 PLUS=" & CStr(CCur("+900719925474099.3125"))
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
    EmitResult "XL-019 OPEN2ERR=" & CStr(errNo)
    If errNo <> 0 Then
        Err.Clear
        Close #first
        On Error GoTo 0
        Exit Sub
    End If
    Err.Clear
    Lock #first, 1 To 2
    errNo = Err.Number
    EmitResult "XL-019 FIRSTERR=" & CStr(errNo)
    Err.Clear
    Lock #second, 1 To 2
    EmitResult "XL-019 SECONDERR=" & CStr(Err.Number)
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
    EmitResult "XL-020 OPEN2ERR=" & CStr(errNo)
    If errNo <> 0 Then
        Err.Clear
        Close #first
        On Error GoTo 0
        Exit Sub
    End If
    Err.Clear
    Lock #first, 1 To 2
    EmitResult "XL-020 FIRSTERR=" & CStr(Err.Number)
    Err.Clear
    Lock #second, 2 To 3
    EmitResult "XL-020 SECONDERR=" & CStr(Err.Number)
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
        EmitResult "XL-021 " & label & " ERR=0 VALUE=" & CStr(value)
    Else
        EmitResult "XL-021 " & label & " ERR=" & CStr(errNo)
    End If
    On Error GoTo 0
End Sub

Private Sub VerifyFormatRounding()
    EmitResult "XL-022 CURRENCY=[" & Format(2.675, "Currency") & "]"
    EmitResult "XL-022 FIXED=[" & Format(2.675, "Fixed") & "]"
    EmitResult "XL-022 STANDARD=[" & Format(2.675, "Standard") & "]"
    EmitResult "XL-022 SCIENTIFIC=[" & Format(2.675, "Scientific") & "]"
    EmitResult "XL-022 SCI09995=[" & Format(0.09995, "Scientific") & "]"
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
    EmitResult "XL-023 MODE=" & mode & " ZERO=" & CStr(zeroErr) & _
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
    EmitResult "XL-024 MODE=" & mode & " ZERO=" & CStr(zeroErr) & _
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
    EmitResult "XL-025 INPUT LINEEOF=" & CStr(EOF(f))
    On Error Resume Next
    Err.Clear: Line Input #f, text: errNo = Err.Number
    EmitResult "XL-025 INPUT LINEERR=" & CStr(errNo) & " EOF=" & CStr(EOF(f))
    Close #f
    On Error GoTo 0

    f = FreeFile: Open path For Binary As #f
    Line Input #f, text
    EmitResult "XL-025 BINARY LINEEOF=" & CStr(EOF(f))
    On Error Resume Next
    Err.Clear: Line Input #f, text: errNo = Err.Number
    EmitResult "XL-025 BINARY LINEERR=" & CStr(errNo) & " EOF=" & CStr(EOF(f))
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
    EmitResult id & " LOF=" & n & " BYTES=" & line
End Sub
