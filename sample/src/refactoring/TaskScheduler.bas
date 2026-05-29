Option Explicit

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
    ' 3. 担当者キャパシティ設定 (AssigneeConfig)
    '    - 設定範囲:
    '      - ROW_START / END で行範囲を指定。
    '      - COL_NAME: 担当者名列。
    '      - COL_LIMIT: キャパシティ上限値列。
    '    - 各担当者の最大工数 (デフォルト 1.0) を設定可能。
    '
    ' 4. マイクロタスク (微小タスク)
    '    - 工数が 0 より大きいが、0.25単位で丸めると 0 になってしまう場合 (例: 0.1, 0.05)。
    '    - 空き容量が 0.1 以上ある日に限り、最低 0.1 の工数を割り当てます (容量不足時はスキップ)。
    '    - 通常タスク (例: 1.01) は 0.25 単位で丸められ、端数は切り捨てられます (例: 1.0)。
    '
    ' 5. ロック (COL_LOCK: "L" Mark)
    '    - "L" マークが付いた行は自動スケジュールの対象外となり、現状のセル値が維持されます。
    '    - ただし、リソース使用量には加算されるため、後続タスクの空き容量計算に影響を与えます。
    '
    ' 6. 休日設定 (ROW_HOLIDAY: 5行目)
    '    - "休" (STR_HOLIDAY_MARK) が設定されている日は、工数の割り当てを行いません (スキップ)。
    '    - 前提: 土日・祝日にはあらかじめ "休" が入力されているものとします。
    '
    ' 7. その他設定 (CalendarConfig / TaskConfig / AssigneeConfig)
    '    - ROW_START: データ開始行 (19行目～)
    '    - COL_CALENDAR_START: カレンダー開始列
    '    - 各種カラムインデックス:
    '      - COL_DURATION: 工数 (Duration)
    '      - COL_ASSIGNEE: 担当者 (Assignee)
    ' =========================================================================================

    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    ' Initialize Configurations
    Dim calCfg As CalendarConfig
    calCfg = InitCalendarConfig()
    Dim taskCfg As TaskConfig
    taskCfg = InitTaskConfig()
    Dim assigneeCfg As AssigneeConfig
    assigneeCfg = InitAssigneeConfig()
    
    ' Performance Settings
    Dim screenUpdateState As Boolean
    Dim calcState As XlCalculation
    Dim eventsState As Boolean
    
    screenUpdateState = Application.ScreenUpdating
    calcState = Application.Calculation
    eventsState = Application.EnableEvents
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    On Error GoTo Cleanup
    
    ' 1. Determine Range Bounds
    Dim lastRow As Long
    lastRow = GetLastTaskRow(ws, taskCfg)
    
    Dim lastCol As Long
    lastCol = GetLastCalendarCol(ws, calCfg)
    
    Dim numRows As Long
    numRows = CalcNumRows(lastRow, taskCfg)
    If numRows < 1 Then GoTo Cleanup
    
    Dim numDays As Long
    numDays = CalcNumDays(lastCol, calCfg)
    If numDays < 1 Then GoTo Cleanup
    
    ' 2. Read Data into Arrays (Double Buffering)
    Dim rangeTask As Range
    Set rangeTask = GetTaskRange(ws, taskCfg, lastRow)
    Dim taskDataFrame As Variant
    taskDataFrame = rangeTask.Value
    
    Dim rangeSchedule As Range
    Set rangeSchedule = GetScheduleRange(ws, taskCfg, calCfg, lastRow, lastCol)
    Dim scheduleGrid As Variant
    scheduleGrid = rangeSchedule.Value
    
    Dim rangeHoliday As Range
    Set rangeHoliday = GetHolidayRange(ws, calCfg, lastCol)
    Dim holidayData As Variant
    holidayData = rangeHoliday.Value 
    
    ' 3. Read Capacity Config
    Dim rangeAssignee As Range
    Set rangeAssignee = GetAssigneeRange(ws, assigneeCfg)
    Dim assigneeDataFrame As Variant
    assigneeDataFrame = rangeAssignee.Value
    
    ' Build Capacity Dict: Key=Name, Value=MaxLimit (Default 1.0)
    Dim capacityLimits As Object
    Set capacityLimits = BuildCapacityDict(assigneeDataFrame)
    
    ' 4. Initialize Resource Usage Dictionary
    Dim assigneeUsage As Object
    Set assigneeUsage = CreateObject("Scripting.Dictionary")
    
    Dim taskRow As Long
    Dim assigneeName As String
    
    ' =========================================================
    ' Phase 1: Scan Locked Rows ("L")
    ' =========================================================
    Call ScanLockedRows(taskCfg, numRows, numDays, taskDataFrame, scheduleGrid, assigneeUsage)
    
    ' =========================================================
    ' Phase 2: Schedule & Calculate Dependencies (Locked & Unlocked)
    ' =========================================================
    ' Level Tracking
    Dim maxLevel As Long
    maxLevel = GetMaxLevel(taskDataFrame, numRows, taskCfg)
    Dim currentLevel As Long
    Dim levelMaxFinish() As Long
    Dim levelMaxFinishAlloc() As Double
    ReDim levelMaxFinish(0 To maxLevel)
    ReDim levelMaxFinishAlloc(0 To maxLevel)
    Dim parentFinishIdx As Long
    Dim parentFinishAlloc As Double
    Dim baseStartIdx As Long
    
    Dim taskFinishIdx As Long
    Dim taskFinishAlloc As Double
    Dim isLocked As Boolean
    
    For taskRow = 1 To numRows
        isLocked = IsRowLocked(taskDataFrame, taskRow, taskCfg)
        assigneeName = GetAssigneeName(taskDataFrame, taskRow, taskCfg)
        
        ' ===================================
        ' 【階層ロジック】: D列に基づく依存関係の計算 (Level 1, 2, 3...)
        ' ===================================
        currentLevel = GetTaskLevel(taskDataFrame, taskRow, taskCfg)
        
        ' Level 1 -> 新しいタスクブロックの開始 (完了日リセット)
        If currentLevel = 1 Then
            Erase levelMaxFinish
            Erase levelMaxFinishAlloc
        End If
        
        ' Level > 1 -> 親レベルの最終日に基づく開始日計算
        parentFinishIdx = 0
        parentFinishAlloc = 0
        If currentLevel > 1 Then
            parentFinishIdx = levelMaxFinish(currentLevel - 1)
            parentFinishAlloc = levelMaxFinishAlloc(currentLevel - 1)
        End If
        baseStartIdx = CalcBaseStartIdx(currentLevel, parentFinishIdx, parentFinishAlloc)
        ' ===================================
        ' 【通常ロジック】: 計算された開始日からスケジュール割り当てへ
        ' ===================================
        
        taskFinishIdx = 0
        taskFinishAlloc = 0
        
        If isLocked Then
            ' Only update dependency state from existing grid
            ' Scan from right to left to find finish
            Call FindLockedTaskFinish(taskRow, numDays, scheduleGrid, taskFinishIdx, taskFinishAlloc)
            
             ' Update Level Max Finish Logic for Locked Row
             Call UpdateLevelFinish(currentLevel, taskFinishIdx, taskFinishAlloc, levelMaxFinish, levelMaxFinishAlloc)
            
        Else
            ' Unlocked -> Schedule it
            Call ScheduleUnlockedTask(taskCfg, calCfg, taskRow, numDays, baseStartIdx, taskDataFrame, holidayData, capacityLimits, scheduleGrid, assigneeUsage, taskFinishIdx, taskFinishAlloc)
            
            ' Update Level Max Finish Logic for Unlocked Row
            Call UpdateLevelFinish(currentLevel, taskFinishIdx, taskFinishAlloc, levelMaxFinish, levelMaxFinishAlloc)
        End If
        
    Next taskRow
    
    ' 4. Write Back to Sheet
    rangeSchedule.Value = scheduleGrid
    
    ' MsgBox "Scheduling Complete!"
    
Cleanup:
    Application.ScreenUpdating = screenUpdateState
    Application.Calculation = calcState
    Application.EnableEvents = eventsState
    
End Sub

