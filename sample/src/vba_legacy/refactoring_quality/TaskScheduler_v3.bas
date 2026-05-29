Option Explicit

' ============================================================
' モジュールレベル定数: COL_ クラスタ（列インデックス）
' ============================================================
Private Const COL_CALENDAR_START_IDX As Long = 24
Private Const COL_LEVEL_IDX          As Long = 8
Private Const COL_OFFSET_IDX         As Long = 9
Private Const COL_LOCK_IDX           As Long = 13
Private Const COL_DURATION_IDX       As Long = 15
Private Const COL_ASSIGNEE_IDX       As Long = 17

' ============================================================
' モジュールレベル定数: ROW_ クラスタ（行インデックス）
' ============================================================
Private Const ROW_HEADER_SHEET  As Long = 3
Private Const ROW_HOLIDAY_SHEET As Long = 5
Private Const ROW_START_SHEET   As Long = 19

' ============================================================
' モジュールレベル定数: CONFIG_ クラスタ（設定範囲）
' ============================================================
Private Const CONFIG_ROW_START As Long = 8
Private Const CONFIG_ROW_END   As Long = 14
Private Const CONFIG_COL_NAME  As Long = 17
Private Const CONFIG_COL_LIMIT As Long = 18

' ============================================================
' モジュールレベル定数: 文字列マーク
' ============================================================
Private Const STR_HOLIDAY_MARK As String = "休"
Private Const STR_LOCK_MARK    As String = "L"

' ============================================================
' モジュールレベル定数: 繰り返し数値リテラルの定数化
' ============================================================
Private Const UNIT_SIZE       As Double = 0.25   ' 工数の最小単位（0.25 人日）
Private Const HALF_UNIT       As Double = 0.5    ' 端数判定に使う 0.5
Private Const COL_FIRST_DATA  As Long   = 1      ' 即値引数: Cells の先頭列インデックス

' ============================================================
' ヘルパー: Application 状態の保存と復元
' ============================================================
Private Sub SaveApplicationState( _
    ByRef outScreenUpdate As Boolean, _
    ByRef outCalc As XlCalculation, _
    ByRef outEvents As Boolean)
    outScreenUpdate = Application.ScreenUpdating
    outCalc         = Application.Calculation
    outEvents       = Application.EnableEvents
End Sub

Private Sub RestoreApplicationState( _
    ByVal screenUpdate As Boolean, _
    ByVal calc As XlCalculation, _
    ByVal events As Boolean)
    Application.ScreenUpdating = screenUpdate
    Application.Calculation    = calc
    Application.EnableEvents   = events
End Sub

' ============================================================
' ヘルパー: シートから 2D 配列を読み込む
'   ws        : 対象シート
'   startRow  : 開始行（シート座標）
'   startCol  : 開始列（シート座標）
'   endRow    : 終了行（シート座標）
'   endCol    : 終了列（シート座標）
'   outRange  : 読み込んだ Range オブジェクト（返値）
'   戻り値    : 配列 Variant
' ============================================================
Private Function ReadRangeToArray( _
    ws As Worksheet, _
    ByVal startRow As Long, ByVal startCol As Long, _
    ByVal endRow   As Long, ByVal endCol   As Long, _
    ByRef outRange As Range) As Variant
    Set outRange    = ws.Range(ws.Cells(startRow, startCol), ws.Cells(endRow, endCol))
    ReadRangeToArray = outRange.Value
End Function

' ============================================================
' ヘルパー: レベル完了日を更新する（重複ブロックを統一）
'   taskFinishIdx   : 今回のタスクの完了日インデックス
'   taskFinishAlloc : 今回の最終日配分量
'   level           : 更新対象のレベル
'   levelMaxFinish  : レベル別最遅完了日配列
'   levelMaxFinishAlloc : レベル別最遅完了日の配分量配列
' ============================================================
Private Sub UpdateLevelMaxFinish( _
    ByVal taskFinishIdx   As Long, _
    ByVal taskFinishAlloc As Double, _
    ByVal level           As Long, _
    ByRef levelMaxFinish()      As Long, _
    ByRef levelMaxFinishAlloc() As Double)
    If taskFinishIdx > levelMaxFinish(level) Then
        levelMaxFinish(level)      = taskFinishIdx
        levelMaxFinishAlloc(level) = taskFinishAlloc
    ElseIf taskFinishIdx = levelMaxFinish(level) Then
        If taskFinishAlloc > levelMaxFinishAlloc(level) Then
            levelMaxFinishAlloc(level) = taskFinishAlloc
        End If
    End If
End Sub

' ============================================================
' ヘルパー: 担当者の使用量配列が未登録なら初期化して登録する
' ============================================================
Private Sub EnsurePersonUsage( _
    ByRef personUsage As Object, _
    ByVal assigneeName As String, _
    ByVal numDays As Long)
    If Not personUsage.Exists(assigneeName) Then
        Dim newAllocArray() As Double
        ReDim newAllocArray(1 To numDays) As Double
        personUsage.Add assigneeName, newAllocArray
    End If
End Sub

' ============================================================
' メインルーチン
' ============================================================
Sub AutoScheduleTasks()
    Dim ws As Worksheet
    Set ws = ActiveSheet

    ' Application 状態の保存
    Dim screenUpdateState As Boolean
    Dim calcState As XlCalculation
    Dim eventsState As Boolean
    Call SaveApplicationState(screenUpdateState, calcState, eventsState)

    Application.ScreenUpdating = False
    Application.Calculation    = xlCalculationManual
    Application.EnableEvents   = False
    On Error GoTo Cleanup

    ' ----------------------------------------------------------
    ' シートの有効範囲を確認する
    ' ----------------------------------------------------------
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, COL_DURATION_IDX).End(xlUp).Row
    If lastRow < ROW_START_SHEET Then GoTo Cleanup

    Dim lastCol As Long
    lastCol = ws.Cells(ROW_HEADER_SHEET, ws.Columns.Count).End(xlToLeft).Column
    If lastCol < COL_CALENDAR_START_IDX Then GoTo Cleanup

    Dim numRows As Long
    numRows = lastRow - ROW_START_SHEET + 1
    Dim numDays As Long
    numDays = lastCol - COL_CALENDAR_START_IDX + 1

    ' ----------------------------------------------------------
    ' シートデータの一括読み込み
    ' ----------------------------------------------------------
    Dim rangeMeta As Range
    Dim metaData As Variant
    metaData = ReadRangeToArray(ws, ROW_START_SHEET, COL_FIRST_DATA, lastRow, CONFIG_COL_NAME, rangeMeta)

    Dim rangeGrid As Range
    Dim gridData As Variant
    gridData = ReadRangeToArray(ws, ROW_START_SHEET, COL_CALENDAR_START_IDX, lastRow, lastCol, rangeGrid)

    Dim rangeHoliday As Range
    Dim holidayData As Variant
    holidayData = ReadRangeToArray(ws, ROW_HOLIDAY_SHEET, COL_CALENDAR_START_IDX, ROW_HOLIDAY_SHEET, lastCol, rangeHoliday)

    Dim rangeConfig As Range
    Dim configData As Variant
    configData = ReadRangeToArray(ws, CONFIG_ROW_START, CONFIG_COL_NAME, CONFIG_ROW_END, CONFIG_COL_LIMIT, rangeConfig)

    ' ----------------------------------------------------------
    ' 担当者ごとの上限負荷を辞書に読み込む
    ' ----------------------------------------------------------
    Dim capacityLimits As Object
    Set capacityLimits = CreateObject("Scripting.Dictionary")
    Dim cfgRow      As Long
    Dim cfgName     As String
    Dim cfgCapacity As Double
    For cfgRow = 1 To UBound(configData, 1)
        cfgName = Trim(configData(cfgRow, 1))
        If cfgName <> "" Then
            cfgCapacity = 1#
            If IsNumeric(configData(cfgRow, 2)) And Not IsEmpty(configData(cfgRow, 2)) Then
                cfgCapacity = CDbl(configData(cfgRow, 2))
            End If
            capacityLimits(cfgName) = cfgCapacity
        End If
    Next cfgRow

    ' ----------------------------------------------------------
    ' 担当者ごとの日別使用量を格納する辞書
    ' ----------------------------------------------------------
    Dim personUsage As Object
    Set personUsage = CreateObject("Scripting.Dictionary")

    Dim taskRow      As Long
    Dim dayIdx       As Long
    Dim assigneeName As String
    Dim cellVal      As Variant
    Dim existingAlloc As Double
    Dim newAllocArray() As Double

    ' ----------------------------------------------------------
    ' Phase 1: ロック行をスキャンして使用量を集計する
    ' ----------------------------------------------------------
    For taskRow = 1 To numRows
        assigneeName = Trim(metaData(taskRow, COL_ASSIGNEE_IDX))
        If assigneeName <> "" Then
            Call EnsurePersonUsage(personUsage, assigneeName, numDays)
            If UCase(Trim(metaData(taskRow, COL_LOCK_IDX))) = STR_LOCK_MARK Then
                newAllocArray = personUsage(assigneeName)
                For dayIdx = 1 To numDays
                    cellVal       = gridData(taskRow, dayIdx)
                    existingAlloc = 0
                    If IsNumeric(cellVal) And Not IsEmpty(cellVal) Then existingAlloc = CDbl(cellVal)
                    If existingAlloc > 0 Then newAllocArray(dayIdx) = newAllocArray(dayIdx) + existingAlloc
                Next dayIdx
                personUsage(assigneeName) = newAllocArray
            End If
        End If
    Next taskRow

    ' ----------------------------------------------------------
    ' Phase 2: タスクをスケジューリングする
    ' ----------------------------------------------------------
    Dim duration         As Double
    Dim remaining        As Double
    Dim dailyAlloc       As Double
    Dim capacity         As Double
    Dim maxDailyLoad     As Double
    Dim isHoliday        As Boolean
    Dim lagDays          As Variant
    Dim taskStartIdx     As Long
    Dim currentLevel     As Long
    Dim levelMaxFinish(0 To 100)      As Long
    Dim levelMaxFinishAlloc(0 To 100) As Double
    Dim parentFinishIdx   As Long
    Dim parentFinishAlloc As Double
    Dim baseStartIdx      As Long
    Dim taskFinishIdx     As Long
    Dim taskFinishAlloc   As Double
    Dim isLocked          As Boolean

    For taskRow = 1 To numRows
        isLocked     = (UCase(Trim(metaData(taskRow, COL_LOCK_IDX))) = STR_LOCK_MARK)
        assigneeName = Trim(metaData(taskRow, COL_ASSIGNEE_IDX))

        ' レベルを取得する
        currentLevel = 0
        If IsNumeric(metaData(taskRow, COL_LEVEL_IDX)) And Not IsEmpty(metaData(taskRow, COL_LEVEL_IDX)) Then
            currentLevel = CLng(metaData(taskRow, COL_LEVEL_IDX))
        End If

        ' レベル 1 に達したらレベル完了日をリセットする
        If currentLevel = 1 Then
            Erase levelMaxFinish
            Erase levelMaxFinishAlloc
        End If

        ' 親レベルの完了日を基準に開始インデックスを決める
        baseStartIdx = 1
        If currentLevel > 1 Then
            parentFinishIdx   = levelMaxFinish(currentLevel - 1)
            parentFinishAlloc = levelMaxFinishAlloc(currentLevel - 1)
            If parentFinishIdx > 0 Then
                If parentFinishAlloc < HALF_UNIT Then
                    baseStartIdx = parentFinishIdx
                Else
                    baseStartIdx = parentFinishIdx + 1
                End If
            End If
        End If

        taskFinishIdx   = 0
        taskFinishAlloc = 0

        If isLocked Then
            ' ロック行: 既存セル値から完了日を探す
            For dayIdx = numDays To 1 Step -1
                cellVal       = gridData(taskRow, dayIdx)
                existingAlloc = 0
                If IsNumeric(cellVal) And Not IsEmpty(cellVal) Then existingAlloc = CDbl(cellVal)
                If existingAlloc > 0 Then
                    taskFinishIdx   = dayIdx
                    taskFinishAlloc = existingAlloc
                    Exit For
                End If
            Next dayIdx
            If currentLevel > 0 Then
                Call UpdateLevelMaxFinish(taskFinishIdx, taskFinishAlloc, currentLevel, levelMaxFinish, levelMaxFinishAlloc)
            End If
        Else
            ' 未ロック行: グリッドをクリアしてから新規割り当てを行う
            duration = 0
            If IsNumeric(metaData(taskRow, COL_DURATION_IDX)) Then duration = CDbl(metaData(taskRow, COL_DURATION_IDX))
            For dayIdx = 1 To numDays: gridData(taskRow, dayIdx) = Empty: Next dayIdx

            If assigneeName <> "" And duration > 0 Then
                Call EnsurePersonUsage(personUsage, assigneeName, numDays)
                maxDailyLoad = 1#
                If capacityLimits.Exists(assigneeName) Then maxDailyLoad = capacityLimits(assigneeName)

                remaining     = duration
                newAllocArray = personUsage(assigneeName)

                ' オフセット（ラグ日数）を加算して開始日を確定する
                taskStartIdx = baseStartIdx
                lagDays = metaData(taskRow, COL_OFFSET_IDX)
                If IsNumeric(lagDays) And Not IsEmpty(lagDays) Then taskStartIdx = taskStartIdx + CLng(lagDays)
                If taskStartIdx < 1 Then taskStartIdx = 1

                ' 0.25 未満の微小タスクを判定する
                Dim totalNeeded As Long
                Dim isMicroTask As Boolean
                totalNeeded = Int((duration / UNIT_SIZE) + HALF_UNIT)
                isMicroTask = (totalNeeded = 0 And duration > 0)

                For dayIdx = taskStartIdx To numDays
                    If remaining <= 0 Then Exit For
                    isHoliday = (Trim(holidayData(1, dayIdx)) = STR_HOLIDAY_MARK)
                    If Not isHoliday Then
                        capacity = maxDailyLoad - newAllocArray(dayIdx)
                        If capacity < 0 Then capacity = 0

                        Dim maxUnits     As Long
                        Dim neededUnits  As Long
                        Dim allocateUnits As Long
                        maxUnits   = Int(capacity / UNIT_SIZE)
                        dailyAlloc = 0

                        If isMicroTask Then
                            ' 微小タスク: 0.1 人日単位で配分する
                            If capacity >= 0.1 Then dailyAlloc = 0.1
                        Else
                            neededUnits = Int((remaining / UNIT_SIZE) + HALF_UNIT)
                            If maxUnits > 0 And neededUnits > 0 Then
                                allocateUnits = neededUnits
                                If allocateUnits > maxUnits Then allocateUnits = maxUnits
                                dailyAlloc = allocateUnits * UNIT_SIZE
                            End If
                        End If

                        If dailyAlloc > 0 Then
                            gridData(taskRow, dayIdx)  = dailyAlloc
                            newAllocArray(dayIdx)      = newAllocArray(dayIdx) + dailyAlloc
                            remaining                  = remaining - dailyAlloc
                            taskFinishIdx              = dayIdx
                            taskFinishAlloc            = dailyAlloc
                        End If
                    End If
                Next dayIdx
                personUsage(assigneeName) = newAllocArray
            End If

            If currentLevel > 0 Then
                Call UpdateLevelMaxFinish(taskFinishIdx, taskFinishAlloc, currentLevel, levelMaxFinish, levelMaxFinishAlloc)
            End If
        End If
    Next taskRow

    ' スケジュール結果をシートに書き戻す
    rangeGrid.Value = gridData

Cleanup:
    ' Application 状態を復元する
    Call RestoreApplicationState(screenUpdateState, calcState, eventsState)
End Sub
