Option Explicit

' =========================================================================================
' TaskScheduler_v5.bas — 自動スケジュールエンジン（リファクタリング済み）
'
' 設計方針:
'   - I/O 層 / ビジネスロジック層 / データ層 の 3 層分離
'   - AutoScheduleTasks は「読む → 計算 → 書く」の I/O シェルのみ
'   - ロジック Function は Excel 非依存（テスト可能）
'   - 各プロシージャ 30 行以内、ネスト 3 以内
' =========================================================================================

' -----------------------------------------------------------------------------------------
' UDT 定義（接頭辞クラスターを型に集約）
' -----------------------------------------------------------------------------------------

' タスクシートのレイアウト定数を束ねる型
Type TaskLayoutConfig
    CalendarStartCol As Long  ' カレンダー開始列（COL_CALENDAR_START_IDX）
    HeaderRow        As Long  ' ヘッダー行（ROW_HEADER_SHEET）
    HolidayRow       As Long  ' 休日行（ROW_HOLIDAY_SHEET）
    LevelCol         As Long  ' レベル列（COL_LEVEL_IDX）
    OffsetCol        As Long  ' オフセット列（COL_OFFSET_IDX）
    LockCol          As Long  ' ロック列（COL_LOCK_IDX）
    DurationCol      As Long  ' 工数列（COL_DURATION_IDX）
    AssigneeCol      As Long  ' 担当者列（COL_ASSIGNEE_IDX）
    DataStartRow     As Long  ' データ開始行（ROW_START_SHEET）
    NameCol          As Long  ' 担当者設定: 名前列（CONFIG_COL_NAME）
End Type

' 担当者設定エリアのレイアウト定数を束ねる型
Type AssigneeLayoutConfig
    RowStart  As Long  ' 設定行開始（CONFIG_ROW_START）
    RowEnd    As Long  ' 設定行終了（CONFIG_ROW_END）
    NameCol   As Long  ' 担当者名列（CONFIG_COL_NAME）
    LimitCol  As Long  ' 上限値列（CONFIG_COL_LIMIT）
End Type

' -----------------------------------------------------------------------------------------
' モジュール定数
' -----------------------------------------------------------------------------------------

' 工数割り当ての基本単位
Private Const ALLOC_UNIT             As Double = 0.25
' 親タスク最終日に「同日開始」と判定する割当閾値
Private Const PARENT_ALLOC_THRESHOLD As Double = 0.5
' マイクロタスクとして割り当てる最低工数
Private Const MICRO_ALLOC_MIN        As Double = 0.1

' 文字列マーク
Private Const STR_HOLIDAY_MARK As String = "休"
Private Const STR_LOCK_MARK    As String = "L"

' -----------------------------------------------------------------------------------------
' ヘルパー：設定値の初期化
' -----------------------------------------------------------------------------------------

' タスクシートレイアウト設定を初期化して返す
Private Function InitTaskLayout() As TaskLayoutConfig
    Dim cfg As TaskLayoutConfig
    cfg.CalendarStartCol = 24
    cfg.HeaderRow        = 3
    cfg.HolidayRow       = 5
    cfg.LevelCol         = 8
    cfg.OffsetCol        = 9
    cfg.LockCol          = 13
    cfg.DurationCol      = 15
    cfg.AssigneeCol      = 17
    cfg.DataStartRow     = 19
    cfg.NameCol          = 17
    InitTaskLayout = cfg
End Function

' 担当者設定エリアのレイアウトを初期化して返す
Private Function InitAssigneeLayout() As AssigneeLayoutConfig
    Dim cfg As AssigneeLayoutConfig
    cfg.RowStart = 8
    cfg.RowEnd   = 14
    cfg.NameCol  = 17
    cfg.LimitCol = 18
    InitAssigneeLayout = cfg
End Function

' -----------------------------------------------------------------------------------------
' ヘルパー：Application 状態の保存・復元
' -----------------------------------------------------------------------------------------

' Application の画面更新・計算・イベントを無効化し、元の状態をオブジェクトで返す
Private Function SaveAppState() As Object
    Dim state As Object
    Set state = CreateObject("Scripting.Dictionary")
    state("ScreenUpdating") = Application.ScreenUpdating
    state("Calculation")    = Application.Calculation
    state("EnableEvents")   = Application.EnableEvents
    Application.ScreenUpdating = False
    Application.Calculation    = xlCalculationManual
    Application.EnableEvents   = False
    Set SaveAppState = state
End Function

' SaveAppState で保存した状態を Application に復元する
Private Sub RestoreAppState(ByVal state As Object)
    Application.ScreenUpdating = state("ScreenUpdating")
    Application.Calculation    = state("Calculation")
    Application.EnableEvents   = state("EnableEvents")
End Sub

' -----------------------------------------------------------------------------------------
' I/O 層：Range → Variant 配列の読み込み
' -----------------------------------------------------------------------------------------

' タスクメタデータ（左端から AssigneeCol 列まで）を Variant 配列で返す
Private Function ReadMetaData(ByVal ws As Worksheet, _
                              ByVal layout As TaskLayoutConfig, _
                              ByVal lastRow As Long) As Variant
    ReadMetaData = ws.Range( _
        ws.Cells(layout.DataStartRow, 1), _
        ws.Cells(lastRow, layout.NameCol)).Value
End Function

' カレンダーグリッドを Variant 配列で返す
Private Function ReadGridData(ByVal ws As Worksheet, _
                              ByVal layout As TaskLayoutConfig, _
                              ByVal lastRow As Long, _
                              ByVal lastCol As Long) As Variant
    ReadGridData = ws.Range( _
        ws.Cells(layout.DataStartRow, layout.CalendarStartCol), _
        ws.Cells(lastRow, lastCol)).Value
End Function

' 休日行を Variant 配列で返す
Private Function ReadHolidayData(ByVal ws As Worksheet, _
                                  ByVal layout As TaskLayoutConfig, _
                                  ByVal lastCol As Long) As Variant
    ReadHolidayData = ws.Range( _
        ws.Cells(layout.HolidayRow, layout.CalendarStartCol), _
        ws.Cells(layout.HolidayRow, lastCol)).Value
End Function

' 担当者設定エリアを Variant 配列で返す
Private Function ReadConfigData(ByVal ws As Worksheet, _
                                ByVal aLayout As AssigneeLayoutConfig) As Variant
    ReadConfigData = ws.Range( _
        ws.Cells(aLayout.RowStart, aLayout.NameCol), _
        ws.Cells(aLayout.RowEnd, aLayout.LimitCol)).Value
End Function

' -----------------------------------------------------------------------------------------
' I/O 層：最終行・最終列の取得
' -----------------------------------------------------------------------------------------

' データの最終行を返す。データが存在しない場合は -1 を返す
Private Function GetLastRow(ByVal ws As Worksheet, _
                             ByVal layout As TaskLayoutConfig) As Long
    Dim r As Long
    r = ws.Cells(ws.Rows.Count, layout.DurationCol).End(xlUp).Row
    If r < layout.DataStartRow Then GetLastRow = -1 Else GetLastRow = r
End Function

' ヘッダー行のカレンダー最終列を返す。カレンダーが存在しない場合は -1 を返す
Private Function GetLastCol(ByVal ws As Worksheet, _
                             ByVal layout As TaskLayoutConfig) As Long
    Dim c As Long
    c = ws.Cells(layout.HeaderRow, ws.Columns.Count).End(xlToLeft).Column
    If c < layout.CalendarStartCol Then GetLastCol = -1 Else GetLastCol = c
End Function

' -----------------------------------------------------------------------------------------
' ビジネスロジック層：担当者キャパシティ辞書の構築
' -----------------------------------------------------------------------------------------

' configData からキャパシティ上限辞書を構築して返す
' configData(row, 1)=名前 / configData(row, 2)=上限値
Private Function BuildCapacityDict(ByRef configData As Variant) As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    Dim r As Long
    For r = 1 To UBound(configData, 1)
        Dim cfgName As String
        cfgName = Trim(configData(r, 1))
        If cfgName <> "" Then
            Dim cap As Double
            cap = 1#
            If IsNumeric(configData(r, 2)) And Not IsEmpty(configData(r, 2)) Then
                cap = CDbl(configData(r, 2))
            End If
            dict(cfgName) = cap
        End If
    Next r

    Set BuildCapacityDict = dict
End Function

' -----------------------------------------------------------------------------------------
' ビジネスロジック層：personUsage 辞書の初期化補助
' -----------------------------------------------------------------------------------------

' personUsage に担当者が未登録なら numDays 分の配列を追加する
Private Sub EnsurePersonUsage(ByVal personUsage As Object, _
                               ByVal assigneeName As String, _
                               ByVal numDays As Long)
    If personUsage.Exists(assigneeName) Then Exit Sub

    Dim arr() As Double
    ReDim arr(1 To numDays) As Double
    personUsage.Add assigneeName, arr
End Sub

' -----------------------------------------------------------------------------------------
' ビジネスロジック層：levelMaxFinish の更新
' -----------------------------------------------------------------------------------------

' タスクの完了日インデックスと完了日割当量で levelMaxFinish 配列を更新する
Private Sub UpdateLevelFinish(ByVal currentLevel As Long, _
                               ByVal taskFinishIdx As Long, _
                               ByVal taskFinishAlloc As Double, _
                               ByRef levelMaxFinish() As Long, _
                               ByRef levelMaxFinishAlloc() As Double)
    If currentLevel <= 0 Then Exit Sub

    If taskFinishIdx > levelMaxFinish(currentLevel) Then
        levelMaxFinish(currentLevel)      = taskFinishIdx
        levelMaxFinishAlloc(currentLevel) = taskFinishAlloc
    ElseIf taskFinishIdx = levelMaxFinish(currentLevel) Then
        If taskFinishAlloc > levelMaxFinishAlloc(currentLevel) Then
            levelMaxFinishAlloc(currentLevel) = taskFinishAlloc
        End If
    End If
End Sub

' -----------------------------------------------------------------------------------------
' ビジネスロジック層：ロック行のスキャン（Phase 1）
' -----------------------------------------------------------------------------------------

' ロック行 1 行分の既存割当を usageArr に積み上げて personUsage に書き戻す
Private Sub AccumulateLockedRow(ByRef gridData As Variant, _
                                 ByVal taskRow As Long, _
                                 ByVal numDays As Long, _
                                 ByVal personUsage As Object, _
                                 ByVal assigneeName As String)
    Dim usageArr() As Double
    usageArr = personUsage(assigneeName)
    Dim dayIdx As Long
    For dayIdx = 1 To numDays
        Dim cellVal As Variant
        cellVal = gridData(taskRow, dayIdx)
        If IsNumeric(cellVal) And Not IsEmpty(cellVal) Then
            Dim alloc As Double
            alloc = CDbl(cellVal)
            If alloc > 0 Then usageArr(dayIdx) = usageArr(dayIdx) + alloc
        End If
    Next dayIdx
    personUsage(assigneeName) = usageArr
End Sub

' ロック行を走査し、各担当者の使用量配列に既存割当を積み上げる
Private Sub ScanLockedRows(ByRef metaData As Variant, _
                            ByRef gridData As Variant, _
                            ByVal numRows As Long, _
                            ByVal numDays As Long, _
                            ByVal layout As TaskLayoutConfig, _
                            ByVal personUsage As Object)
    Dim taskRow As Long
    For taskRow = 1 To numRows
        Dim assigneeName As String
        assigneeName = Trim(metaData(taskRow, layout.AssigneeCol))
        If assigneeName <> "" Then
            EnsurePersonUsage personUsage, assigneeName, numDays
            If UCase(Trim(metaData(taskRow, layout.LockCol))) = STR_LOCK_MARK Then
                AccumulateLockedRow gridData, taskRow, numDays, personUsage, assigneeName
            End If
        End If
    Next taskRow
End Sub

' -----------------------------------------------------------------------------------------
' ビジネスロジック層：1 日分の割当量を計算する
' -----------------------------------------------------------------------------------------

' 1 日分の割当量を計算して返す。割当不可の場合は 0 を返す
Private Function CalcDailyAlloc(ByVal remaining As Double, _
                                 ByVal capacity As Double, _
                                 ByVal isMicroTask As Boolean) As Double
    If isMicroTask Then
        If capacity >= MICRO_ALLOC_MIN Then CalcDailyAlloc = MICRO_ALLOC_MIN
        Exit Function
    End If

    Dim maxUnits    As Long
    Dim neededUnits As Long
    maxUnits    = Int(capacity / ALLOC_UNIT)
    neededUnits = Int((remaining / ALLOC_UNIT) + 0.5)
    If maxUnits <= 0 Or neededUnits <= 0 Then Exit Function

    Dim allocUnits As Long
    allocUnits = neededUnits
    If allocUnits > maxUnits Then allocUnits = maxUnits
    CalcDailyAlloc = allocUnits * ALLOC_UNIT
End Function

' -----------------------------------------------------------------------------------------
' ビジネスロジック層：ロック行の完了日情報を確定する
' -----------------------------------------------------------------------------------------

' ロック行に対して gridData を後ろから走査し taskFinishIdx / taskFinishAlloc を確定する
Private Sub ResolveLockedTaskFinish(ByRef gridData As Variant, _
                                     ByVal taskRow As Long, _
                                     ByVal numDays As Long, _
                                     ByRef taskFinishIdx As Long, _
                                     ByRef taskFinishAlloc As Double)
    taskFinishIdx   = 0
    taskFinishAlloc = 0
    Dim dayIdx As Long
    For dayIdx = numDays To 1 Step -1
        Dim cellVal As Variant
        cellVal = gridData(taskRow, dayIdx)
        If IsNumeric(cellVal) And Not IsEmpty(cellVal) Then
            Dim alloc As Double
            alloc = CDbl(cellVal)
            If alloc > 0 Then
                taskFinishIdx   = dayIdx
                taskFinishAlloc = alloc
                Exit For
            End If
        End If
    Next dayIdx
End Sub

' -----------------------------------------------------------------------------------------
' ビジネスロジック層：開始日インデックスを計算する
' -----------------------------------------------------------------------------------------

' 親レベルの完了情報から基本開始日インデックスを求め、ラグ加算後の値を返す
Private Function CalcTaskStartIdx(ByVal currentLevel As Long, _
                                   ByRef lagRaw As Variant, _
                                   ByRef levelMaxFinish() As Long, _
                                   ByRef levelMaxFinishAlloc() As Double) As Long
    Dim baseIdx As Long
    baseIdx = 1

    If currentLevel > 1 Then
        Dim parentIdx   As Long
        Dim parentAlloc As Double
        parentIdx   = levelMaxFinish(currentLevel - 1)
        parentAlloc = levelMaxFinishAlloc(currentLevel - 1)
        If parentIdx > 0 Then
            If parentAlloc < PARENT_ALLOC_THRESHOLD Then
                baseIdx = parentIdx
            Else
                baseIdx = parentIdx + 1
            End If
        End If
    End If

    If IsNumeric(lagRaw) And Not IsEmpty(lagRaw) Then
        baseIdx = baseIdx + CLng(lagRaw)
    End If
    If baseIdx < 1 Then baseIdx = 1
    CalcTaskStartIdx = baseIdx
End Function

' -----------------------------------------------------------------------------------------
' ビジネスロジック層：非ロック行のスケジューリング（日次割り当てループ）
' -----------------------------------------------------------------------------------------

' 指定タスクの日次割り当てループを実行する
' 戻り値: 最終割り当て日インデックス（未割当なら 0）/ taskFinishAlloc は ByRef で返す
Private Function AllocateDays(ByRef gridData As Variant, _
                               ByRef holidayData As Variant, _
                               ByVal taskRow As Long, _
                               ByVal numDays As Long, _
                               ByVal taskStartIdx As Long, _
                               ByVal maxDailyLoad As Double, _
                               ByVal duration As Double, _
                               ByRef usageArr() As Double, _
                               ByRef taskFinishAlloc As Double) As Long
    Dim isMicroTask As Boolean
    isMicroTask = (Int((duration / ALLOC_UNIT) + 0.5) = 0 And duration > 0)

    Dim remaining As Double
    remaining = duration

    Dim dayIdx As Long
    For dayIdx = taskStartIdx To numDays
        If Trim(holidayData(1, dayIdx)) <> STR_HOLIDAY_MARK Then
            Dim capacity As Double
            capacity = maxDailyLoad - usageArr(dayIdx)
            If capacity < 0 Then capacity = 0

            Dim dailyAlloc As Double
            dailyAlloc = CalcDailyAlloc(remaining, capacity, isMicroTask)
            If dailyAlloc > 0 Then
                gridData(taskRow, dayIdx) = dailyAlloc
                usageArr(dayIdx)          = usageArr(dayIdx) + dailyAlloc
                remaining                 = remaining - dailyAlloc
                AllocateDays              = dayIdx
                taskFinishAlloc           = dailyAlloc
            End If
        End If
        If remaining <= 0 Then Exit For  ' 残工数消化でループ終了
    Next dayIdx
End Function

' グリッド行をクリアし、工数と担当者を確認してから AllocateDays を呼び出す
' 戻り値: 最終割り当て日インデックス（0 = スキップ）/ taskFinishAlloc は ByRef
Private Function ScheduleOneTask(ByRef metaData As Variant, _
                                  ByRef gridData As Variant, _
                                  ByRef holidayData As Variant, _
                                  ByVal taskRow As Long, _
                                  ByVal numDays As Long, _
                                  ByVal taskStartIdx As Long, _
                                  ByVal maxDailyLoad As Double, _
                                  ByVal personUsage As Object, _
                                  ByVal assigneeName As String, _
                                  ByRef taskFinishAlloc As Double) As Long
    Dim duration As Double
    Dim col15 As Variant
    col15 = metaData(taskRow, 15) ' DurationCol は 15 固定
    If IsNumeric(col15) Then duration = CDbl(col15)

    Dim dayIdx As Long
    For dayIdx = 1 To numDays  ' グリッド行をクリア
        gridData(taskRow, dayIdx) = Empty
    Next dayIdx

    ScheduleOneTask = 0
    taskFinishAlloc = 0
    If assigneeName = "" Or duration <= 0 Then Exit Function

    EnsurePersonUsage personUsage, assigneeName, numDays
    Dim usageArr() As Double
    usageArr = personUsage(assigneeName)

    ScheduleOneTask = AllocateDays(gridData, holidayData, taskRow, numDays, _
                                   taskStartIdx, maxDailyLoad, duration, _
                                   usageArr, taskFinishAlloc)
    personUsage(assigneeName) = usageArr
End Function

' -----------------------------------------------------------------------------------------
' ビジネスロジック層：非ロック行のディスパッチ
' -----------------------------------------------------------------------------------------

' 非ロック 1 タスク行の開始日・最大負荷を求め ScheduleOneTask を呼び出す
' 戻り値: taskFinishIdx / taskFinishAlloc は ByRef
Private Function DispatchUnlockedTask(ByRef metaData As Variant, _
                                       ByRef gridData As Variant, _
                                       ByRef holidayData As Variant, _
                                       ByVal taskRow As Long, _
                                       ByVal numDays As Long, _
                                       ByVal currentLevel As Long, _
                                       ByVal layout As TaskLayoutConfig, _
                                       ByVal capacityLimits As Object, _
                                       ByVal personUsage As Object, _
                                       ByRef levelMaxFinish() As Long, _
                                       ByRef levelMaxFinishAlloc() As Double, _
                                       ByRef taskFinishAlloc As Double) As Long
    Dim assigneeName As String
    assigneeName = Trim(metaData(taskRow, layout.AssigneeCol))

    Dim taskStartIdx As Long
    taskStartIdx = CalcTaskStartIdx(currentLevel, _
                                    metaData(taskRow, layout.OffsetCol), _
                                    levelMaxFinish, levelMaxFinishAlloc)

    Dim maxLoad As Double
    maxLoad = 1#
    If capacityLimits.Exists(assigneeName) Then maxLoad = capacityLimits(assigneeName)

    DispatchUnlockedTask = ScheduleOneTask(metaData, gridData, holidayData, _
                                           taskRow, numDays, taskStartIdx, maxLoad, _
                                           personUsage, assigneeName, taskFinishAlloc)
End Function

' -----------------------------------------------------------------------------------------
' ビジネスロジック層：スケジューリングループ（Phase 2）
' -----------------------------------------------------------------------------------------

' 全タスク行を走査してグリッドにスケジュールを書き込む（メモリ内操作のみ）
Private Sub ScheduleAllTasks(ByRef metaData As Variant, _
                              ByRef gridData As Variant, _
                              ByRef holidayData As Variant, _
                              ByVal numRows As Long, _
                              ByVal numDays As Long, _
                              ByVal layout As TaskLayoutConfig, _
                              ByVal capacityLimits As Object, _
                              ByVal personUsage As Object)
    Dim levelMaxFinish(0 To 100)      As Long
    Dim levelMaxFinishAlloc(0 To 100) As Double

    Dim taskRow As Long
    For taskRow = 1 To numRows
        Dim isLocked As Boolean
        isLocked = (UCase(Trim(metaData(taskRow, layout.LockCol))) = STR_LOCK_MARK)

        Dim currentLevel As Long
        currentLevel = 0
        If IsNumeric(metaData(taskRow, layout.LevelCol)) And _
           Not IsEmpty(metaData(taskRow, layout.LevelCol)) Then
            currentLevel = CLng(metaData(taskRow, layout.LevelCol))
        End If
        If currentLevel = 1 Then
            Erase levelMaxFinish
            Erase levelMaxFinishAlloc
        End If

        Dim taskFinishIdx   As Long
        Dim taskFinishAlloc As Double
        If isLocked Then
            Call ResolveLockedTaskFinish(gridData, taskRow, numDays, _
                                        taskFinishIdx, taskFinishAlloc)
        Else
            taskFinishIdx = DispatchUnlockedTask( _
                metaData, gridData, holidayData, taskRow, numDays, currentLevel, _
                layout, capacityLimits, personUsage, _
                levelMaxFinish, levelMaxFinishAlloc, taskFinishAlloc)
        End If
        UpdateLevelFinish currentLevel, taskFinishIdx, taskFinishAlloc, _
                          levelMaxFinish, levelMaxFinishAlloc
    Next taskRow
End Sub

' -----------------------------------------------------------------------------------------
' エントリーポイント（I/O シェル）
' -----------------------------------------------------------------------------------------

' シートからデータを読み込んでビジネスロジックに渡し、結果を書き戻す
' 引数: ws=対象シート / tLayout, aLayout=レイアウト設定 / lastRow, lastCol=範囲
Private Sub RunScheduleCore(ByVal ws As Worksheet, _
                             ByVal tLayout As TaskLayoutConfig, _
                             ByVal aLayout As AssigneeLayoutConfig, _
                             ByVal lastRow As Long, _
                             ByVal lastCol As Long)
    Dim numRows As Long
    Dim numDays As Long
    numRows = lastRow - tLayout.DataStartRow + 1
    numDays = lastCol - tLayout.CalendarStartCol + 1

    Dim metaData    As Variant
    Dim gridData    As Variant
    Dim holidayData As Variant
    Dim configData  As Variant
    metaData    = ReadMetaData(ws, tLayout, lastRow)
    gridData    = ReadGridData(ws, tLayout, lastRow, lastCol)
    holidayData = ReadHolidayData(ws, tLayout, lastCol)
    configData  = ReadConfigData(ws, aLayout)

    Dim capacityLimits As Object
    Set capacityLimits = BuildCapacityDict(configData)
    Dim personUsage As Object
    Set personUsage = CreateObject("Scripting.Dictionary")

    ScanLockedRows metaData, gridData, numRows, numDays, tLayout, personUsage
    ScheduleAllTasks metaData, gridData, holidayData, numRows, numDays, _
                     tLayout, capacityLimits, personUsage

    ws.Range( _
        ws.Cells(tLayout.DataStartRow, tLayout.CalendarStartCol), _
        ws.Cells(lastRow, lastCol)).Value = gridData
End Sub

' 自動スケジュール マクロのエントリーポイント
' 設計: 本 Sub は I/O シェル。スケジューリングは RunScheduleCore に委譲する
Sub AutoScheduleTasks()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.ActiveSheet  ' マクロ実行時にアクティブなシートを対象とする

    Dim tLayout As TaskLayoutConfig
    tLayout = InitTaskLayout()
    Dim aLayout As AssigneeLayoutConfig
    aLayout = InitAssigneeLayout()

    Dim appState As Object
    Set appState = SaveAppState()
    On Error GoTo Cleanup

    Dim lastRow As Long
    lastRow = GetLastRow(ws, tLayout)
    If lastRow = -1 Then GoTo Cleanup

    Dim lastCol As Long
    lastCol = GetLastCol(ws, tLayout)
    If lastCol = -1 Then GoTo Cleanup

    RunScheduleCore ws, tLayout, aLayout, lastRow, lastCol

Cleanup:
    RestoreAppState appState
End Sub
