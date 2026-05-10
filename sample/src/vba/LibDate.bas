Option Explicit

' ============================================================
' LibDate.bas — 日付ユーティリティ関数ライブラリ
'
' VBA 標準ライブラリを補完するよく使う日付ヘルパー関数群。
' Excel オブジェクトには依存せず単体テスト可能。
' ============================================================

' --------------------------------------------------------------
' 判定系
' --------------------------------------------------------------

' 閏年かどうかを返す
Function IsLeapYear(y As Integer) As Boolean
    IsLeapYear = (y Mod 4 = 0 And y Mod 100 <> 0) Or (y Mod 400 = 0)
End Function

' 土曜または日曜かどうかを返す
Function IsWeekend(d As Date) As Boolean
    Dim w As Integer
    w = Weekday(d)
    IsWeekend = (w = vbSunday Or w = vbSaturday)
End Function

' 月〜金かどうかを返す
Function IsWeekdayDate(d As Date) As Boolean
    IsWeekdayDate = Not IsWeekend(d)
End Function

' d が startDate 以上 endDate 以下かどうかを返す（両端含む）
Function IsBetweenDates(d As Date, startDate As Date, endDate As Date) As Boolean
    IsBetweenDates = (d >= startDate And d <= endDate)
End Function

' --------------------------------------------------------------
' 月・四半期・年の端点
' --------------------------------------------------------------

' 指定した年月の日数を返す
Function DaysInMonth(y As Integer, m As Integer) As Integer
    DaysInMonth = Day(DateSerial(y, m + 1, 0))
End Function

' 月の初日を返す
Function StartOfMonth(d As Date) As Date
    StartOfMonth = DateSerial(Year(d), Month(d), 1)
End Function

' 月の末日を返す
Function EndOfMonth(d As Date) As Date
    EndOfMonth = DateSerial(Year(d), Month(d) + 1, 0)
End Function

' 四半期番号を返す (1〜4)
Function QuarterOf(d As Date) As Integer
    QuarterOf = (Month(d) - 1) \ 3 + 1
End Function

' 四半期の初日を返す
Function StartOfQuarter(d As Date) As Date
    Dim q As Integer
    q = QuarterOf(d)
    StartOfQuarter = DateSerial(Year(d), (q - 1) * 3 + 1, 1)
End Function

' 四半期の末日を返す
Function EndOfQuarter(d As Date) As Date
    Dim q As Integer
    q = QuarterOf(d)
    EndOfQuarter = DateSerial(Year(d), q * 3 + 1, 0)
End Function

' 年の初日を返す
Function StartOfYear(d As Date) As Date
    StartOfYear = DateSerial(Year(d), 1, 1)
End Function

' 年の末日を返す
Function EndOfYear(d As Date) As Date
    EndOfYear = DateSerial(Year(d), 12, 31)
End Function

' --------------------------------------------------------------
' 曜日検索
' --------------------------------------------------------------

' d 以降（d を含む）で最初に dow 曜日が現れる日付を返す
' dow: vbMonday〜vbSunday
Function NextOrSameWeekday(d As Date, dow As Integer) As Date
    Dim diff As Integer
    diff = (dow - Weekday(d) + 7) Mod 7
    NextOrSameWeekday = d + diff
End Function

' d より後で最初に dow 曜日が現れる日付を返す（d 自身は含まない）
Function NextWeekday(d As Date, dow As Integer) As Date
    Dim diff As Integer
    diff = (dow - Weekday(d) + 7) Mod 7
    If diff = 0 Then diff = 7
    NextWeekday = d + diff
End Function

' d 以前（d を含む）で最後に dow 曜日が現れる日付を返す
Function PrevOrSameWeekday(d As Date, dow As Integer) As Date
    Dim diff As Integer
    diff = (Weekday(d) - dow + 7) Mod 7
    PrevOrSameWeekday = d - diff
End Function

' 指定月の第 n 週の dow 曜日を返す（例: 第3月曜日）
' n=1〜5。n=5 で存在しない場合は第4週を返す。
Function NthWeekdayOfMonth(y As Integer, m As Integer, n As Integer, dow As Integer) As Date
    Dim firstDay As Date
    firstDay = DateSerial(y, m, 1)
    Dim firstDow As Date
    firstDow = NextOrSameWeekday(firstDay, dow)
    Dim result As Date
    result = firstDow + (n - 1) * 7
    ' n=5 で翌月にはみ出した場合は 1週 戻す
    If Month(result) <> m Then result = result - 7
    NthWeekdayOfMonth = result
End Function

' 指定月の最終 dow 曜日を返す（例: 最終金曜日）
Function LastWeekdayOfMonth(y As Integer, m As Integer, dow As Integer) As Date
    LastWeekdayOfMonth = PrevOrSameWeekday(DateSerial(y, m + 1, 0), dow)
End Function

' --------------------------------------------------------------
' 営業日計算（祝日は考慮しない）
' --------------------------------------------------------------

' d1（含む）から d2（含まない）までの営業日数（土日除く）を返す
' d2 < d1 の場合は負値を返す
Function WorkdaysBetween(d1 As Date, d2 As Date) As Long
    Dim totalDays As Long
    totalDays = DateDiff("d", d1, d2)
    If totalDays = 0 Then
        WorkdaysBetween = 0
        Exit Function
    End If

    Dim sign As Integer
    sign = IIf(totalDays > 0, 1, -1)
    totalDays = Abs(totalDays)

    ' 完全な週数 × 5 + 端数の平日数
    Dim startD As Date
    If sign = 1 Then startD = d1 Else startD = d2
    ' 開始日の週内位置（月=0）
    Dim startDow As Integer
    startDow = (Weekday(startD) + 5) Mod 7   ' Mon=0..Fri=4,Sat=5,Sun=6

    Dim fullWeeks As Long
    Dim remainder As Long
    fullWeeks = totalDays \ 7
    remainder = totalDays Mod 7

    Dim weekdays As Long
    weekdays = fullWeeks * 5

    ' 端数分の平日を 1 日ずつ数える
    Dim i As Long
    For i = 0 To remainder - 1
        Dim pos As Integer
        pos = (startDow + i) Mod 7
        If pos < 5 Then weekdays = weekdays + 1   ' Mon〜Fri
    Next i

    WorkdaysBetween = sign * weekdays
End Function

' n 営業日後（負なら前）の日付を返す
Function AddWorkdays(d As Date, n As Long) As Date
    If n = 0 Then
        AddWorkdays = d
        Exit Function
    End If

    Dim dir As Integer
    dir = IIf(n > 0, 1, -1)
    Dim remaining As Long
    remaining = Abs(n)
    Dim cur As Date
    cur = d

    Do While remaining > 0
        cur = cur + dir
        If Not IsWeekend(cur) Then remaining = remaining - 1
    Loop

    AddWorkdays = cur
End Function

' --------------------------------------------------------------
' ISO 8601 週番号
' --------------------------------------------------------------

' ISO 8601 週番号を返す（月曜始まり、1月4日を含む週が第1週）
Function ISOWeekNumber(d As Date) As Integer
    ' d が属する週の木曜日を求める
    Dim dow As Integer
    dow = (Weekday(d) + 5) Mod 7      ' Mon=0..Sun=6
    Dim thursday As Date
    thursday = d + (3 - dow)

    ' その年の第1週の木曜日
    Dim jan4 As Date
    jan4 = DateSerial(Year(thursday), 1, 4)
    Dim jan4dow As Integer
    jan4dow = (Weekday(jan4) + 5) Mod 7
    Dim firstThursday As Date
    firstThursday = jan4 + (3 - jan4dow)

    ISOWeekNumber = DateDiff("d", firstThursday, thursday) \ 7 + 1
End Function

' --------------------------------------------------------------
' 書式変換
' --------------------------------------------------------------

' 日付を "YYYY-MM-DD" 形式の文字列に返す
Function ToDateString(d As Date) As String
    Dim m As String
    Dim dv As String
    m  = Right("0" & Month(d), 2)
    dv = Right("0" & Day(d), 2)
    ToDateString = Year(d) & "-" & m & "-" & dv
End Function

' "YYYY-MM-DD" 形式の文字列を日付に変換して返す
Function ParseDateString(s As String) As Date
    Dim y As Integer, mon As Integer, dv As Integer
    y   = CInt(Left(s, 4))
    mon = CInt(Mid(s, 6, 2))
    dv  = CInt(Right(s, 2))
    ParseDateString = DateSerial(y, mon, dv)
End Function

' 日付を "YYYYMMDD" 形式の 8 桁文字列で返す
Function ToDateKey(d As Date) As String
    Dim m As String
    Dim dv As String
    m  = Right("0" & Month(d), 2)
    dv = Right("0" & Day(d), 2)
    ToDateKey = Year(d) & m & dv
End Function

' --------------------------------------------------------------
' 年齢・経過年数
' --------------------------------------------------------------

' refDate 時点の満年齢（年数）を返す
Function AgeInYears(birthDate As Date, refDate As Date) As Integer
    Dim age As Integer
    age = Year(refDate) - Year(birthDate)
    Dim notYetBirthday As Boolean
    notYetBirthday = (Month(refDate) < Month(birthDate)) Or (Month(refDate) = Month(birthDate) And Day(refDate) < Day(birthDate))
    If notYetBirthday Then age = age - 1
    AgeInYears = age
End Function
