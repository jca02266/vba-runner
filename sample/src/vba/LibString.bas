Option Explicit

' ============================================================
' LibString.bas — 文字列ユーティリティ関数ライブラリ
'
' VBA 標準ライブラリを補完するよく使う文字列ヘルパー関数群。
' Excel オブジェクトには依存せず単体テスト可能。
' ============================================================

' --------------------------------------------------------------
' 判定系
' --------------------------------------------------------------

' s が空文字列または空白のみかどうかを返す
Function IsNullOrEmpty(s As String) As Boolean
    IsNullOrEmpty = (Len(Trim(s)) = 0)
End Function

' s の全文字が数字（0〜9）かどうかを返す
Function IsAllDigits(s As String) As Boolean
    If Len(s) = 0 Then
        IsAllDigits = False
        Exit Function
    End If
    Dim i As Long
    Dim c As String
    For i = 1 To Len(s)
        c = Mid(s, i, 1)
        If c < "0" Or c > "9" Then
            IsAllDigits = False
            Exit Function
        End If
    Next i
    IsAllDigits = True
End Function

' s の全文字がアルファベット（a〜z または A〜Z）かどうかを返す
Function IsAllAlpha(s As String) As Boolean
    If Len(s) = 0 Then
        IsAllAlpha = False
        Exit Function
    End If
    Dim i As Long
    Dim c As String
    For i = 1 To Len(s)
        c = LCase(Mid(s, i, 1))
        If c < "a" Or c > "z" Then
            IsAllAlpha = False
            Exit Function
        End If
    Next i
    IsAllAlpha = True
End Function

' s が prefix で始まるかどうかを返す（大文字小文字を区別）
Function StartsWith(s As String, prefix As String) As Boolean
    StartsWith = (Left(s, Len(prefix)) = prefix)
End Function

' s が suffix で終わるかどうかを返す（大文字小文字を区別）
Function EndsWith(s As String, suffix As String) As Boolean
    EndsWith = (Right(s, Len(suffix)) = suffix)
End Function

' s に substr が含まれるかどうかを返す（大文字小文字を区別）
Function Contains(s As String, substr As String) As Boolean
    Contains = (InStr(s, substr) > 0)
End Function

' --------------------------------------------------------------
' 検索・カウント
' --------------------------------------------------------------

' s の中に substr が何回現れるかを返す（重複なし）
Function CountOccurrences(s As String, substr As String) As Long
    Dim count As Long
    Dim pos As Long
    Dim subLen As Long
    subLen = Len(substr)
    If subLen = 0 Then
        CountOccurrences = 0
        Exit Function
    End If
    pos = 1
    Do
        pos = InStr(pos, s, substr)
        If pos = 0 Then Exit Do
        count = count + 1
        pos = pos + subLen
    Loop
    CountOccurrences = count
End Function

' --------------------------------------------------------------
' 変換・整形
' --------------------------------------------------------------

' s の左側を padChar で埋めて最低 n 文字にして返す
Function PadLeft(s As String, n As Integer, padChar As String) As String
    Dim result As String
    result = s
    Do While Len(result) < n
        result = Left(padChar, 1) & result
    Loop
    PadLeft = result
End Function

' s の右側を padChar で埋めて最低 n 文字にして返す
Function PadRight(s As String, n As Integer, padChar As String) As String
    Dim result As String
    result = s
    Do While Len(result) < n
        result = result & Left(padChar, 1)
    Loop
    PadRight = result
End Function

' 数値 n を width 桁にゼロ埋めした文字列を返す
Function ZeroPad(n As Long, width As Integer) As String
    ZeroPad = PadLeft(CStr(n), width, "0")
End Function

' 先頭・末尾の空白を除去し、連続する空白を1つにまとめて返す
Function TrimAll(s As String) As String
    Dim result As String
    result = Trim(s)
    Do While InStr(result, "  ") > 0
        result = Replace(result, "  ", " ")
    Loop
    TrimAll = result
End Function

' 最初の文字を大文字、残りを小文字にして返す
Function Capitalize(s As String) As String
    If Len(s) = 0 Then
        Capitalize = s
        Exit Function
    End If
    Capitalize = UCase(Left(s, 1)) & LCase(Mid(s, 2))
End Function

' s を n 回繰り返した文字列を返す
Function RepeatStr(s As String, n As Long) As String
    Dim result As String
    Dim i As Long
    For i = 1 To n
        result = result & s
    Next i
    RepeatStr = result
End Function

' s の文字を逆順にした文字列を返す
Function ReverseString(s As String) As String
    Dim result As String
    Dim i As Long
    For i = Len(s) To 1 Step -1
        result = result & Mid(s, i, 1)
    Next i
    ReverseString = result
End Function

' s が maxLen 文字を超えていれば末尾を ellipsis に置き換えて返す
Function Truncate(s As String, maxLen As Integer, ellipsis As String) As String
    If Len(s) <= maxLen Then
        Truncate = s
    Else
        Truncate = Left(s, maxLen - Len(ellipsis)) & ellipsis
    End If
End Function

' --------------------------------------------------------------
' 分割・抽出
' --------------------------------------------------------------

' 最初の delimiter より前の部分を返す（見つからなければ s 全体を返す）
Function SubstringBefore(s As String, delimiter As String) As String
    Dim pos As Long
    pos = InStr(s, delimiter)
    If pos = 0 Then
        SubstringBefore = s
    Else
        SubstringBefore = Left(s, pos - 1)
    End If
End Function

' 最初の delimiter より後ろの部分を返す（見つからなければ "" を返す）
Function SubstringAfter(s As String, delimiter As String) As String
    Dim pos As Long
    pos = InStr(s, delimiter)
    If pos = 0 Then
        SubstringAfter = ""
    Else
        SubstringAfter = Mid(s, pos + Len(delimiter))
    End If
End Function

' 最後の delimiter より前の部分を返す（見つからなければ s 全体を返す）
Function SubstringBeforeLast(s As String, delimiter As String) As String
    Dim pos As Long
    pos = InStrRev(s, delimiter)
    If pos = 0 Then
        SubstringBeforeLast = s
    Else
        SubstringBeforeLast = Left(s, pos - 1)
    End If
End Function

' 最後の delimiter より後ろの部分を返す（見つからなければ "" を返す）
Function SubstringAfterLast(s As String, delimiter As String) As String
    Dim pos As Long
    pos = InStrRev(s, delimiter)
    If pos = 0 Then
        SubstringAfterLast = ""
    Else
        SubstringAfterLast = Mid(s, pos + Len(delimiter))
    End If
End Function

' --------------------------------------------------------------
' 集計
' --------------------------------------------------------------

' 空白区切りの単語数を返す
Function WordCount(s As String) As Long
    Dim trimmed As String
    trimmed = TrimAll(s)
    If Len(trimmed) = 0 Then
        WordCount = 0
        Exit Function
    End If
    Dim parts() As String
    parts = Split(trimmed, " ")
    WordCount = UBound(parts) - LBound(parts) + 1
End Function
