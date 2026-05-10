Option Explicit

' ============================================================
' LibPath.bas — パス名操作ユーティリティ関数ライブラリ
'
' Windows パス（バックスラッシュ区切り）を前提とした
' ファイルパス操作ヘルパー関数群。
' Excel オブジェクトには依存せず単体テスト可能。
' ============================================================

' パス区切り文字を返す
Function PathSep() As String
    PathSep = "\"
End Function

' パス区切りを \ に統一する（/ を \ に変換）
Function NormalizeSeparators(p As String) As String
    NormalizeSeparators = Replace(p, "/", "\")
End Function

' --------------------------------------------------------------
' 分解
' --------------------------------------------------------------

' パスからファイル名（拡張子付き）を返す
' 例: "C:\foo\bar.txt" → "bar.txt"
Function GetFileName(p As String) As String
    Dim norm As String
    norm = NormalizeSeparators(p)
    Dim pos As Long
    pos = InStrRev(norm, "\")
    If pos = 0 Then
        GetFileName = norm
    Else
        GetFileName = Mid(norm, pos + 1)
    End If
End Function

' パスから拡張子（ドット含む）を返す。なければ "" を返す
' 例: "C:\foo\bar.txt" → ".txt"
Function GetExtension(p As String) As String
    Dim fname As String
    fname = GetFileName(p)
    Dim pos As Long
    pos = InStrRev(fname, ".")
    If pos = 0 Then
        GetExtension = ""
    Else
        GetExtension = Mid(fname, pos)
    End If
End Function

' パスから拡張子なしのファイル名を返す
' 例: "C:\foo\bar.txt" → "bar"
Function GetBaseName(p As String) As String
    Dim fname As String
    fname = GetFileName(p)
    Dim pos As Long
    pos = InStrRev(fname, ".")
    If pos = 0 Then
        GetBaseName = fname
    Else
        GetBaseName = Left(fname, pos - 1)
    End If
End Function

' パスの親ディレクトリを返す（末尾の区切り文字は除去）
' 例: "C:\foo\bar" → "C:\foo"
' 例: "C:\foo" → "C:"
Function GetDirectory(p As String) As String
    Dim norm As String
    norm = NormalizeSeparators(p)
    ' 末尾の区切り文字を除去（ルート "C:\" の区切りは残す）
    Do While Len(norm) > 3 And Right(norm, 1) = "\"
        norm = Left(norm, Len(norm) - 1)
    Loop
    Dim pos As Long
    pos = InStrRev(norm, "\")
    If pos = 0 Then
        GetDirectory = ""
    ElseIf pos = 1 Then
        GetDirectory = "\"
    ElseIf pos = 3 And Mid(norm, 2, 1) = ":" Then
        ' "C:\foo" → "C:\" (ドライブルートは区切り文字を残す)
        GetDirectory = Left(norm, 3)
    Else
        GetDirectory = Left(norm, pos - 1)
    End If
End Function

' ドライブレター（例: "C"）を返す。なければ "" を返す
Function GetDriveLetter(p As String) As String
    If Len(p) >= 2 And Mid(p, 2, 1) = ":" Then
        GetDriveLetter = Left(p, 1)
    Else
        GetDriveLetter = ""
    End If
End Function

' --------------------------------------------------------------
' 判定
' --------------------------------------------------------------

' 絶対パス（ドライブレターまたは UNC \\）かどうかを返す
Function IsAbsolutePath(p As String) As Boolean
    Dim norm As String
    norm = NormalizeSeparators(p)
    If Left(norm, 2) = "\\" Then
        IsAbsolutePath = True
        Exit Function
    End If
    If Len(norm) >= 2 And Mid(norm, 2, 1) = ":" Then
        IsAbsolutePath = True
        Exit Function
    End If
    IsAbsolutePath = False
End Function

' UNC パス（\\server\share 形式）かどうかを返す
Function IsUNCPath(p As String) As Boolean
    IsUNCPath = (Left(NormalizeSeparators(p), 2) = "\\")
End Function

' --------------------------------------------------------------
' 結合・変換
' --------------------------------------------------------------

' ディレクトリとファイル名を結合してパスを作る
' 二重の区切り文字が生じないよう調整する
' 例: PathJoin("C:\foo\", "bar.txt") → "C:\foo\bar.txt"
Function PathJoin(dirPath As String, fname As String) As String
    Dim d As String
    d = NormalizeSeparators(dirPath)
    Dim f As String
    f = NormalizeSeparators(fname)
    ' dir の末尾の \ を除去
    Do While Len(d) > 0 And Right(d, 1) = "\"
        d = Left(d, Len(d) - 1)
    Loop
    ' fname の先頭の \ を除去
    Do While Len(f) > 0 And Left(f, 1) = "\"
        f = Mid(f, 2)
    Loop
    If Len(d) = 0 Then
        PathJoin = f
    ElseIf Len(f) = 0 Then
        PathJoin = d
    Else
        PathJoin = d & "\" & f
    End If
End Function

' 末尾に区切り文字 \ がなければ追加して返す
Function EnsureTrailingSlash(p As String) As String
    Dim norm As String
    norm = NormalizeSeparators(p)
    If Len(norm) = 0 Then
        EnsureTrailingSlash = ""
    ElseIf Right(norm, 1) <> "\" Then
        EnsureTrailingSlash = norm & "\"
    Else
        EnsureTrailingSlash = norm
    End If
End Function

' 末尾の区切り文字 \ があれば除去して返す（ルートは除く）
Function RemoveTrailingSlash(p As String) As String
    Dim norm As String
    norm = NormalizeSeparators(p)
    ' "C:\" "\" のようなルートは変更しない
    If Len(norm) <= 3 Then
        RemoveTrailingSlash = norm
        Exit Function
    End If
    If Right(norm, 1) = "\" Then
        RemoveTrailingSlash = Left(norm, Len(norm) - 1)
    Else
        RemoveTrailingSlash = norm
    End If
End Function

' ファイルの拡張子を newExt に置き換えたパスを返す
' newExt にはドットを含める（例: ".csv"）
' 例: ChangeExtension("C:\foo\bar.txt", ".csv") → "C:\foo\bar.csv"
Function ChangeExtension(p As String, newExt As String) As String
    Dim dir As String
    dir = GetDirectory(p)
    Dim base As String
    base = GetBaseName(p)
    If Len(dir) = 0 Then
        ChangeExtension = base & newExt
    Else
        ChangeExtension = EnsureTrailingSlash(dir) & base & newExt
    End If
End Function

' --------------------------------------------------------------
' 深さ
' --------------------------------------------------------------

' パスのディレクトリ深さ（区切り文字の数）を返す
' 例: "C:\foo\bar" → 2, "C:\foo" → 1, "C:\" → 0
Function PathDepth(p As String) As Long
    Dim norm As String
    norm = NormalizeSeparators(p)
    ' 末尾の \ を除去（"C:\" は "C:" になる）
    Do While Len(norm) > 1 And Right(norm, 1) = "\"
        norm = Left(norm, Len(norm) - 1)
    Loop
    ' \ の個数を数える
    Dim count As Long
    Dim i As Long
    For i = 1 To Len(norm)
        If Mid(norm, i, 1) = "\" Then count = count + 1
    Next i
    ' ドライブレター "C:" の : の後の \ は深さに含めない → count はそのまま
    ' "C:\foo\bar" の \ は 2個 → depth = 2 (OK)
    ' ただしドライブルート "C:" は \ が 0個 → depth = 0 (OK)
    ' UNC "\\server\share" の \ は 3個 → 先頭の \\ 2個を差し引く
    If Left(norm, 2) = "\\" Then count = count - 2
    PathDepth = count
End Function
