Option Explicit

' ===========================================================================
' 定数定義
' ===========================================================================

' カレンダー列のインデックス（A列=1）
Private Const COL_CALENDAR_START_IDX As Long = 24

' シート行インデックス
Private Const ROW_HEADER_SHEET    As Long = 3
Private Const ROW_HOLIDAY_SHEET   As Long = 5
Private Const ROW_START_SHEET     As Long = 19

' メタデータ列インデックス（A列=1）
Private Const COL_LEVEL_IDX       As Long = 8  ' 階層レベル
Private Const COL_OFFSET_IDX      As Long = 9  ' 開始オフセット (Lag日数)
Private Const COL_LOCK_IDX        As Long = 13 ' ロックマーク
Private Const COL_DURATION_IDX    As Long = 15 ' 工数 (Duration)
Private Const COL_ASSIGNEE_IDX    As Long = 17 ' 担当者

' キャパシティ設定範囲
Private Const CONFIG_ROW_START    As Long = 8
Private Const CONFIG_ROW_END      As Long = 14
Private Const CONFIG_COL_NAME     As Long = 17 ' 担当者名列 (I列)
Private Const CONFIG_COL_LIMIT    As Long = 18 ' 上限値列 (J列)

' 特殊マーク文字列
Private Const STR_HOLIDAY_MARK    As String = "休"
Private Const STR_LOCK_MARK       As String = "L"

' 工数の最小単位
Private Const ALLOC_UNIT          As Double = 0.25
Private Const MICRO_TASK_MIN      As Double = 0.1

' ===========================================================================
' エントリーポイント
' ===========================================================================

Sub AutoScheduleTasks()
    ' =========================================================================================
    ' 【仕様説明】 - 自動スケジュールロジック v2.0
    '
    ' 1. 階層タスクと依存関係 (Level)
    '    - Level 1: 新しいタスクブロックの開始。これまでの依存関係をリセットします。
    '    - Level N: 直近の親タスク (Level N-1) の完了状況に依存して開始日が決まります。
    '
    ' 2. 開始日の判定 (Allocation Logic & Start Offset)
    '    - 基本開始日: 親タスク(Level N-1)の最終日翌日。
    '      ただし、親タスク最終日の工数割当 < 0.5 の場合は【同日開始】となります。
    '    - 相対開始日 (Start Offset):
    '      基本開始日に Offset の値 (Lag日数) を加算します。
    '      例: 親完了翌日(基本) + 2日(Lag) = 3日後開始
    '
    ' 3. 担当者キャパシティ設定 (Configurable via Constants)
    '    - 設定範囲: I8:J12 (デフォルト)
    '      - CONFIG_ROW_START / END で行範囲を指定。
    '      - CONFIG_COL_NAME: 担当者名列。
    '      - CONFIG_COL_LIMIT: キャパシティ上限値列。
    '    - 各担当者の最大工数 (デフォルト 1.0) を設定可能。
    '
    ' 4. マイクロタスク (微小タスク)
    '    - 工数が 0 より大きいが、0.25単位で丸めると 0 になってしまう場合 (例: 0.1, 0.05)。
    '    - 空き容量が 0.1 以上ある日に限り、最低 0.1 の工数を割り当てます (容量不足時はスキップ)。
    '    - 通常タスク (例: 1.01) は 0.25 単位で丸められ、端数は切り捨てられます (例: 1.0)。
    '
    ' 5. ロック (COL_LOCK_IDX: "L" Mark)
    '    - "L" マークが付いた行は自動スケジュールの対象外となり、現状のセル値が維持されます。
    '    - ただし、リソース使用量には加算されるため、後続タスクの空き容量計算に影響を与えます。
    '
    ' 6. 休日設定 (ROW_HOLIDAY_SHEET: 5行目)
    '    - "休" (STR_HOLIDAY_MARK) が設定されている日は、工数の割り当てを行いません (スキップ)。
    '    - 前提: 土日・祝日にはあらかじめ "休" が入力されているものとします。
    '
    ' 7. その他設定 (Constants)
    '    - ROW_START_SHEET: データ開始行 (19行目～)
    '    - COL_CALENDAR_START_IDX: カレンダー開始列 (X列=24列目)
    '    - 各種カラムインデックス:
    '      - COL_DURATION_IDX: 工数 (Duration)
    '      - COL_ASSIGNEE_IDX: 担当者 (Assignee)
    ' =========================================================================================

    Dim ws As Worksheet
    Set ws = ActiveSheet

    ' パフォーマンス設定の退避と変更
    Dim savedState As TAppState
    SaveAppState savedState
    SetAppStateFast

    On Error GoTo Cleanup

    ' 範囲境界の確定
    Dim bounds As TSheetBounds
    If Not TryGetSheetBounds(ws, bounds) Then GoTo Cleanup

    ' シートデータの読み込み（二重バッファリング）
    Dim metaData As Variant
    Dim gridData As Variant
    Dim holidayData As Variant
    ReadSheetData ws, bounds, metaData, gridData, holidayData

    ' キャパシティ設定の読み込み
    Dim capacityLimits As Object
    Set capacityLimits = BuildCapacityLimits(ws)

    ' 担当者ごとのリソース使用量ディクショナリの初期化
    Dim personUsage As Object
    Set personUsage = CreateObject("Scripting.Dictionary")

    ' フェーズ 1: ロック行のリソース使用量を集計する
    ScanLockedRows metaData, gridData, personUsage, bounds.numRows, bounds.numDays

    ' フェーズ 2: 未ロック行のスケジューリング
    ScheduleAllTasks metaData, gridData, holidayData, personUsage, capacityLimits, bounds.numRows, bounds.numDays

    ' 計算結果をシートに書き戻す
    ws.Range(ws.Cells(ROW_START_SHEET, COL_CALENDAR_START_IDX), _
             ws.Cells(ROW_START_SHEET + bounds.numRows - 1, COL_CALENDAR_START_IDX + bounds.numDays - 1)) _
      .Value = gridData

Cleanup:
    RestoreAppState savedState
End Sub

' ===========================================================================
' アプリケーション状態の管理
' ===========================================================================

' アプリケーション状態を保持する型
Private Type TAppState
    ScreenUpdating As Boolean
    Calculation    As XlCalculation
    EnableEvents   As Boolean
End Type

' 現在のアプリケーション状態を退避する
Private Sub SaveAppState(ByRef state As TAppState)
    state.ScreenUpdating = Application.ScreenUpdating
    state.Calculation    = Application.Calculation
    state.EnableEvents   = Application.EnableEvents
End Sub

' 画面更新・イベント・自動計算を無効化して処理を高速化する
Private Sub SetAppStateFast()
    Application.ScreenUpdating = False
    Application.Calculation    = xlCalculationManual
    Application.EnableEvents   = False
End Sub

' 退避したアプリケーション状態を復元する
Private Sub RestoreAppState(ByRef state As TAppState)
    Application.ScreenUpdating = state.ScreenUpdating
    Application.Calculation    = state.Calculation
    Application.EnableEvents   = state.EnableEvents
End Sub

' ===========================================================================
' シート範囲の確定とデータ読み込み
' ===========================================================================

' シートの有効範囲を保持する型
Private Type TSheetBounds
    lastRow  As Long
    lastCol  As Long
    numRows  As Long
    numDays  As Long
End Type

' データ範囲の境界を取得する。有効範囲が存在する場合は True を返す
Private Function TryGetSheetBounds(ByVal ws As Worksheet, ByRef bounds As TSheetBounds) As Boolean
    TryGetSheetBounds = False

    bounds.lastRow = ws.Cells(ws.Rows.Count, COL_DURATION_IDX).End(xlUp).Row
    If bounds.lastRow < ROW_START_SHEET Then Exit Function

    bounds.lastCol = ws.Cells(ROW_HEADER_SHEET, ws.Columns.Count).End(xlToLeft).Column
    If bounds.lastCol < COL_CALENDAR_START_IDX Then Exit Function

    bounds.numRows = bounds.lastRow - ROW_START_SHEET + 1
    bounds.numDays = bounds.lastCol - COL_CALENDAR_START_IDX + 1

    TryGetSheetBounds = True
End Function

' メタデータ・グリッドデータ・休日データをシートから配列に読み込む
Private Sub ReadSheetData(ByVal ws As Worksheet, ByVal bounds As TSheetBounds, _
                          ByRef metaData As Variant, _
                          ByRef gridData As Variant, _
                          ByRef holidayData As Variant)
    metaData = ws.Range(ws.Cells(ROW_START_SHEET, 1), _
                        ws.Cells(bounds.lastRow, CONFIG_COL_NAME)).Value

    gridData = ws.Range(ws.Cells(ROW_START_SHEET, COL_CALENDAR_START_IDX), _
                        ws.Cells(bounds.lastRow, bounds.lastCol)).Value

    holidayData = ws.Range(ws.Cells(ROW_HOLIDAY_SHEET, COL_CALENDAR_START_IDX), _
                           ws.Cells(ROW_HOLIDAY_SHEET, bounds.lastCol)).Value
End Sub

' ===========================================================================
' キャパシティ設定の読み込み
' ===========================================================================

' 担当者キャパシティ設定（I8:J14）を読み込み、Dictionary（Key=名前, Value=上限）を返す
Private Function BuildCapacityLimits(ByVal ws As Worksheet) As Object
    Dim configData As Variant
    configData = ws.Range(ws.Cells(CONFIG_ROW_START, CONFIG_COL_NAME), _
                          ws.Cells(CONFIG_ROW_END, CONFIG_COL_LIMIT)).Value

    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")

    Dim cfgRow As Long
    For cfgRow = 1 To UBound(configData, 1)
        Dim cfgName As String
        cfgName = Trim(configData(cfgRow, 1))
        If cfgName <> "" Then
            Dim cfgCapacity As Double
            cfgCapacity = 1# ' デフォルト上限
            If IsNumeric(configData(cfgRow, 2)) And Not IsEmpty(configData(cfgRow, 2)) Then
                cfgCapacity = CDbl(configData(cfgRow, 2))
            End If
            result(cfgName) = cfgCapacity
        End If
    Next cfgRow

    Set BuildCapacityLimits = result
End Function

' ===========================================================================
' フェーズ 1: ロック行のリソース集計
' ===========================================================================

' ロックされた行の既存工数を personUsage に加算する
' ロック行は自動スケジュールの対象外だが、後続タスクの空き容量計算に影響を与える
Private Sub ScanLockedRows(ByRef metaData As Variant, ByRef gridData As Variant, _
                           ByVal personUsage As Object, _
                           ByVal numRows As Long, ByVal numDays As Long)
    Dim taskRow As Long
    For taskRow = 1 To numRows
        If IsLockedRow(metaData, taskRow) Then
            Dim assigneeName As String
            assigneeName = Trim(metaData(taskRow, COL_ASSIGNEE_IDX))
            If assigneeName <> "" Then
                EnsurePersonExists personUsage, assigneeName, numDays

                Dim allocArray() As Double
                allocArray = personUsage(assigneeName)

                Dim dayIdx As Long
                For dayIdx = 1 To numDays
                    Dim alloc As Double
                    alloc = ToDouble(gridData(taskRow, dayIdx))
                    If alloc > 0 Then
                        allocArray(dayIdx) = allocArray(dayIdx) + alloc
                    End If
                Next dayIdx

                personUsage(assigneeName) = allocArray
            End If
        End If
    Next taskRow
End Sub

' ===========================================================================
' フェーズ 2: 全タスクのスケジューリング
' ===========================================================================

' 全行を順番に処理し、ロック行はレベル追跡のみ、未ロック行はスケジューリングを行う
Private Sub ScheduleAllTasks(ByRef metaData As Variant, ByRef gridData As Variant, _
                             ByRef holidayData As Variant, ByVal personUsage As Object, _
                             ByVal capacityLimits As Object, _
                             ByVal numRows As Long, ByVal numDays As Long)
    ' レベルごとの最終割り当て日インデックスと最終日工数
    Dim levelMaxFinish(0 To 100)      As Long
    Dim levelMaxFinishAlloc(0 To 100) As Double

    Dim taskRow As Long
    For taskRow = 1 To numRows
        Dim currentLevel As Long
        currentLevel = ToLevel(metaData(taskRow, COL_LEVEL_IDX))

        ' Level 1 は新しいブロックの開始なのでレベル追跡をリセットする
        If currentLevel = 1 Then
            Erase levelMaxFinish
            Erase levelMaxFinishAlloc
        End If

        ' 親タスクの完了状況から基本開始日を決定する
        Dim baseStartIdx As Long
        baseStartIdx = CalcBaseStartIdx(currentLevel, levelMaxFinish, levelMaxFinishAlloc)

        Dim taskFinishIdx   As Long
        Dim taskFinishAlloc As Double

        If IsLockedRow(metaData, taskRow) Then
            ' ロック行: スケジューリングはスキップ、最終割り当て日だけ把握する
            FindLockedTaskFinish gridData, taskRow, numDays, taskFinishIdx, taskFinishAlloc
        Else
            ' 未ロック行: スケジューリングを実行する
            ScheduleTask metaData, gridData, holidayData, personUsage, capacityLimits, _
                         taskRow, numDays, baseStartIdx, taskFinishIdx, taskFinishAlloc
        End If

        ' レベル別の最終完了日を更新する
        UpdateLevelFinish currentLevel, taskFinishIdx, taskFinishAlloc, _
                          levelMaxFinish, levelMaxFinishAlloc
    Next taskRow
End Sub

' ===========================================================================
' 依存関係の計算
' ===========================================================================

' 親タスク（Level N-1）の完了状況から基本開始日インデックスを返す
' 親の最終日工数 < 0.5 なら同日開始、そうでなければ翌日開始
Private Function CalcBaseStartIdx(ByVal currentLevel As Long, _
                                  ByRef levelMaxFinish()      As Long, _
                                  ByRef levelMaxFinishAlloc() As Double) As Long
    CalcBaseStartIdx = 1
    If currentLevel <= 1 Then Exit Function

    Dim parentFinishIdx   As Long
    Dim parentFinishAlloc As Double
    parentFinishIdx   = levelMaxFinish(currentLevel - 1)
    parentFinishAlloc = levelMaxFinishAlloc(currentLevel - 1)

    If parentFinishIdx <= 0 Then Exit Function

    If parentFinishAlloc < 0.5 Then
        CalcBaseStartIdx = parentFinishIdx       ' 同日開始
    Else
        CalcBaseStartIdx = parentFinishIdx + 1   ' 翌日開始
    End If
End Function

' レベル別の最終完了日と最終日工数を更新する
Private Sub UpdateLevelFinish(ByVal currentLevel As Long, _
                              ByVal taskFinishIdx As Long, ByVal taskFinishAlloc As Double, _
                              ByRef levelMaxFinish()      As Long, _
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

' ===========================================================================
' 個別タスクのスケジューリング
' ===========================================================================

' ロック行の最終割り当て日インデックスと工数を後ろから探して返す
Private Sub FindLockedTaskFinish(ByRef gridData As Variant, ByVal taskRow As Long, ByVal numDays As Long, _
                                 ByRef taskFinishIdx As Long, ByRef taskFinishAlloc As Double)
    taskFinishIdx   = 0
    taskFinishAlloc = 0

    Dim dayIdx As Long
    For dayIdx = numDays To 1 Step -1
        Dim alloc As Double
        alloc = ToDouble(gridData(taskRow, dayIdx))
        If alloc > 0 Then
            taskFinishIdx   = dayIdx
            taskFinishAlloc = alloc
            Exit For
        End If
    Next dayIdx
End Sub

' 未ロック行に対してスケジューリングを行い、gridData を更新する
' taskFinishIdx / taskFinishAlloc に最終割り当て日情報を返す
Private Sub ScheduleTask(ByRef metaData As Variant, ByRef gridData As Variant, _
                         ByRef holidayData As Variant, ByVal personUsage As Object, _
                         ByVal capacityLimits As Object, _
                         ByVal taskRow As Long, ByVal numDays As Long, ByVal baseStartIdx As Long, _
                         ByRef taskFinishIdx As Long, ByRef taskFinishAlloc As Double)
    taskFinishIdx   = 0
    taskFinishAlloc = 0

    Dim duration     As Double
    Dim assigneeName As String
    duration     = ToDouble(metaData(taskRow, COL_DURATION_IDX))
    assigneeName = Trim(metaData(taskRow, COL_ASSIGNEE_IDX))

    ' 担当者・工数がない行はスキップ
    If assigneeName = "" Or duration <= 0 Then Exit Sub

    EnsurePersonExists personUsage, assigneeName, numDays

    ' 担当者のキャパシティ上限を取得する（未登録の場合はデフォルト 1.0）
    Dim maxDailyLoad As Double
    maxDailyLoad = 1#
    If capacityLimits.Exists(assigneeName) Then
        maxDailyLoad = capacityLimits(assigneeName)
    End If

    ' 開始日インデックスにオフセット (Lag) を加算する
    Dim taskStartIdx As Long
    taskStartIdx = baseStartIdx
    Dim lagDays As Variant
    lagDays = metaData(taskRow, COL_OFFSET_IDX)
    If IsNumeric(lagDays) And Not IsEmpty(lagDays) Then
        taskStartIdx = taskStartIdx + CLng(lagDays)
    End If
    If taskStartIdx < 1 Then taskStartIdx = 1

    ' マイクロタスク判定（0.25 単位で丸めると 0 になるほど小さい工数）
    Dim isMicroTask As Boolean
    isMicroTask = (Int(duration / ALLOC_UNIT + 0.5) = 0 And duration > 0)

    ' グリッドの該当行をクリアしてから残工数を消化するまで日付を順に処理する
    Dim dayIdx       As Long
    Dim remaining    As Double
    Dim allocArray() As Double
    Dim dailyAlloc   As Double

    allocArray = personUsage(assigneeName)
    remaining  = duration

    For dayIdx = 1 To numDays
        gridData(taskRow, dayIdx) = Empty
    Next dayIdx

    For dayIdx = taskStartIdx To numDays
        If remaining <= 0 Then Exit For

        ' 休日以外の日に工数を割り当てる
        If Trim(holidayData(1, dayIdx)) <> STR_HOLIDAY_MARK Then
            dailyAlloc = CalcDailyAlloc(remaining, allocArray(dayIdx), maxDailyLoad, isMicroTask)

            If dailyAlloc > 0 Then
                gridData(taskRow, dayIdx)  = dailyAlloc
                allocArray(dayIdx)         = allocArray(dayIdx) + dailyAlloc
                remaining                  = remaining - dailyAlloc
                taskFinishIdx              = dayIdx
                taskFinishAlloc            = dailyAlloc
            End If
        End If
    Next dayIdx

    personUsage(assigneeName) = allocArray
End Sub

' 1日分の割り当て工数を計算して返す
' マイクロタスクは空き容量 0.1 以上なら 0.1 を割り当て、
' 通常タスクは 0.25 単位に丸めて割り当てる
Private Function CalcDailyAlloc(ByVal remaining As Double, ByVal usedToday As Double, _
                                ByVal maxDailyLoad As Double, ByVal isMicroTask As Boolean) As Double
    Dim capacity As Double
    capacity = maxDailyLoad - usedToday
    If capacity < 0 Then capacity = 0

    If isMicroTask Then
        If capacity >= MICRO_TASK_MIN Then
            CalcDailyAlloc = MICRO_TASK_MIN
        Else
            CalcDailyAlloc = 0
        End If
        Exit Function
    End If

    Dim maxUnits    As Long
    Dim neededUnits As Long
    Dim allocUnits  As Long
    maxUnits    = Int(capacity / ALLOC_UNIT)
    neededUnits = Int(remaining / ALLOC_UNIT + 0.5)

    If maxUnits > 0 And neededUnits > 0 Then
        allocUnits     = neededUnits
        If allocUnits > maxUnits Then allocUnits = maxUnits
        CalcDailyAlloc = allocUnits * ALLOC_UNIT
    Else
        CalcDailyAlloc = 0
    End If
End Function

' ===========================================================================
' ユーティリティ関数
' ===========================================================================

' 行がロックされているか判定する
Private Function IsLockedRow(ByRef metaData As Variant, ByVal taskRow As Long) As Boolean
    IsLockedRow = (UCase(Trim(metaData(taskRow, COL_LOCK_IDX))) = STR_LOCK_MARK)
End Function

' Variant を Double に安全に変換する。変換不可の場合は 0 を返す
Private Function ToDouble(ByVal val As Variant) As Double
    If IsNumeric(val) And Not IsEmpty(val) Then
        ToDouble = CDbl(val)
    Else
        ToDouble = 0
    End If
End Function

' Variant から階層レベル値を取得する。無効な場合は 0 を返す
Private Function ToLevel(ByVal val As Variant) As Long
    If IsNumeric(val) And Not IsEmpty(val) Then
        ToLevel = CLng(val)
    Else
        ToLevel = 0
    End If
End Function

' personUsage に指定担当者のエントリーがなければ初期化して追加する
Private Sub EnsurePersonExists(ByVal personUsage As Object, ByVal assigneeName As String, ByVal numDays As Long)
    If Not personUsage.Exists(assigneeName) Then
        Dim newArray() As Double
        ReDim newArray(1 To numDays)
        personUsage.Add assigneeName, newArray
    End If
End Sub
