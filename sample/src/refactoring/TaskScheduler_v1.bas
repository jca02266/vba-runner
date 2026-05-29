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
    '    - COL_CALENDAR_START_IDX: カレンダー開始列 (P列=16列目)
    '    - 各種カラムインデックス:
    '      - COL_DURATION_IDX: 工数 (Duration)
    '      - COL_ASSIGNEE_IDX: 担当者 (Assignee)
    ' =========================================================================================

    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    ' Constants
    Const COL_CALENDAR_START_IDX As Long = 24

    Const ROW_HEADER_SHEET As Long = 3
    Const ROW_HOLIDAY_SHEET As Long = 5
    
    ' Column Mappings (1-based relative to A=1)
    Const COL_LEVEL_IDX As Long = 8 ' Level
    Const COL_OFFSET_IDX As Long = 9 ' Start Offset
    Const COL_LOCK_IDX As Long = 13 ' Lock
    Const COL_DURATION_IDX As Long = 15 ' Duration
    Const COL_ASSIGNEE_IDX As Long = 17 ' Assignee
    
    Const STR_HOLIDAY_MARK As String = "休"
    Const STR_LOCK_MARK As String = "L"
    
    ' Config Areas
    Const CONFIG_ROW_START As Long = 8
    Const CONFIG_ROW_END As Long = 14
    Const CONFIG_COL_NAME As Long = 17 ' I
    Const CONFIG_COL_LIMIT As Long = 18 ' J
    Const ROW_START_SHEET As Long = 19
    
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
    lastRow = ws.Cells(ws.Rows.Count, COL_DURATION_IDX).End(xlUp).Row
    If lastRow < ROW_START_SHEET Then GoTo Cleanup
    
    Dim lastCol As Long
    lastCol = ws.Cells(ROW_HEADER_SHEET, ws.Columns.Count).End(xlToLeft).Column
    If lastCol < COL_CALENDAR_START_IDX Then GoTo Cleanup
    
    Dim numRows As Long
    numRows = lastRow - ROW_START_SHEET + 1
    Dim numDays As Long
    numDays = lastCol - COL_CALENDAR_START_IDX + 1
    
    ' 2. Read Data into Arrays (Double Buffering)
    Dim rangeMeta As Range
    Set rangeMeta = ws.Range(ws.Cells(ROW_START_SHEET, 1), ws.Cells(lastRow, CONFIG_COL_NAME))
    Dim metaData As Variant
    metaData = rangeMeta.Value
    
    Dim rangeGrid As Range
    Set rangeGrid = ws.Range(ws.Cells(ROW_START_SHEET, COL_CALENDAR_START_IDX), ws.Cells(lastRow, lastCol))
    Dim gridData As Variant
    gridData = rangeGrid.Value
    
    Dim rangeHoliday As Range
    Set rangeHoliday = ws.Range(ws.Cells(ROW_HOLIDAY_SHEET, COL_CALENDAR_START_IDX), ws.Cells(ROW_HOLIDAY_SHEET, lastCol))
    Dim holidayData As Variant
    holidayData = rangeHoliday.Value 
    
    ' 3. Read Capacity Config (I8:J12)
    Dim rangeConfig As Range
    Set rangeConfig = ws.Range(ws.Cells(CONFIG_ROW_START, CONFIG_COL_NAME), ws.Cells(CONFIG_ROW_END, CONFIG_COL_LIMIT))
    Dim configData As Variant
    configData = rangeConfig.Value
    
    ' Build Capacity Dict: Key=Name, Value=MaxLimit (Default 1.0)
    Dim capacityLimits As Object
    Set capacityLimits = CreateObject("Scripting.Dictionary")
    
    Dim cfgRow As Long
    Dim cfgCapacity As Double
    Dim cfgName As String
    For cfgRow = 1 To UBound(configData, 1)
        cfgName = Trim(configData(cfgRow, 1)) ' Column I
        If cfgName <> "" Then
            cfgCapacity = 1# ' Default
            If IsNumeric(configData(cfgRow, 2)) And Not IsEmpty(configData(cfgRow, 2)) Then
                cfgCapacity = CDbl(configData(cfgRow, 2))
            End If
            capacityLimits(cfgName) = cfgCapacity
        End If
    Next cfgRow
    
    ' 4. Initialize Resource Usage Dictionary
    Dim personUsage As Object
    Set personUsage = CreateObject("Scripting.Dictionary")
    
    Dim taskRow As Long, dayIdx As Long
    Dim assigneeName As String
    Dim newAllocArray() As Double
    Dim cellVal As Variant
    Dim existingAlloc As Double
    
    ' =========================================================
    ' Phase 1: Scan Locked Rows ("L")
    ' =========================================================
    For taskRow = 1 To numRows
        assigneeName = Trim(metaData(taskRow, COL_ASSIGNEE_IDX))
        
        If assigneeName <> "" Then
            If Not personUsage.Exists(assigneeName) Then
                ReDim newAllocArray(1 To numDays) As Double
                personUsage.Add assigneeName, newAllocArray
            End If
            
            If UCase(Trim(metaData(taskRow, COL_LOCK_IDX))) = STR_LOCK_MARK Then
                newAllocArray = personUsage(assigneeName)
                For dayIdx = 1 To numDays
                    cellVal = gridData(taskRow, dayIdx)
                    existingAlloc = 0
                    If IsNumeric(cellVal) And Not IsEmpty(cellVal) Then
                        existingAlloc = CDbl(cellVal)
                    End If
                    
                    If existingAlloc > 0 Then
                        newAllocArray(dayIdx) = newAllocArray(dayIdx) + existingAlloc
                    End If
                Next dayIdx
                personUsage(assigneeName) = newAllocArray
            End If
        End If
    Next taskRow
    
    ' =========================================================
    ' Phase 2: Schedule & Calculate Dependencies (Locked & Unlocked)
    ' =========================================================
    Dim duration As Double
    Dim remaining As Double
    Dim dailyAlloc As Double
    Dim capacity As Double
    Dim maxDailyLoad As Double
    Dim isHoliday As Boolean
    Dim lagDays As Variant
    Dim taskStartIdx As Long
    Dim prevRowIdx As Long
    Dim prevStartIdx As Long
    
    ' Level Tracking
    Dim currentLevel As Long
    Dim levelMaxFinish(0 To 100) As Long
    Dim levelMaxFinishAlloc(0 To 100) As Double
    Dim parentFinishIdx As Long
    Dim parentFinishAlloc As Double
    Dim baseStartIdx As Long
    
    Dim taskFinishIdx As Long
    Dim taskFinishAlloc As Double
    Dim isLocked As Boolean
    
    For taskRow = 1 To numRows
        isLocked = (UCase(Trim(metaData(taskRow, COL_LOCK_IDX))) = STR_LOCK_MARK)
        assigneeName = Trim(metaData(taskRow, COL_ASSIGNEE_IDX))
        
        ' ===================================
        ' 【階層ロジック】: D列に基づく依存関係の計算 (Level 1, 2, 3...)
        ' ===================================
        currentLevel = 0
        If IsNumeric(metaData(taskRow, COL_LEVEL_IDX)) And Not IsEmpty(metaData(taskRow, COL_LEVEL_IDX)) Then
            currentLevel = CLng(metaData(taskRow, COL_LEVEL_IDX))
        End If
        
        ' Level 1 -> 新しいタスクブロックの開始 (完了日リセット)
        If currentLevel = 1 Then
            Erase levelMaxFinish
            Erase levelMaxFinishAlloc
        End If
        
        ' Level > 1 -> 親レベルの最終日に基づく開始日計算
        baseStartIdx = 1
        If currentLevel > 1 Then
            parentFinishIdx = levelMaxFinish(currentLevel - 1)
            parentFinishAlloc = levelMaxFinishAlloc(currentLevel - 1)
            
            If parentFinishIdx > 0 Then
                ' ロジック: 前レベルのタスクの最終日の割り当てが0.5未満の場合、
                ' 依存タスクは同日に開始可能です。
                ' それ以外（0.5以上）の場合、翌日に開始する必要があります。
                If parentFinishAlloc < 0.5 Then
                    baseStartIdx = parentFinishIdx
                Else
                    baseStartIdx = parentFinishIdx + 1
                End If
            End If
        End If
        ' ===================================
        ' 【通常ロジック】: 計算された開始日からスケジュール割り当てへ
        ' ===================================
        
        taskFinishIdx = 0
        taskFinishAlloc = 0
        
        If isLocked Then
            ' Only update dependency state from existing grid
            ' Scan from right to left to find finish
            For dayIdx = numDays To 1 Step -1
                cellVal = gridData(taskRow, dayIdx)
                existingAlloc = 0
                If IsNumeric(cellVal) And Not IsEmpty(cellVal) Then existingAlloc = CDbl(cellVal)
                
                If existingAlloc > 0 Then
                    taskFinishIdx = dayIdx
                    taskFinishAlloc = existingAlloc
                    Exit For
                End If
            Next dayIdx
            
             ' Update Level Max Finish Logic for Locked Row
            If currentLevel > 0 Then
                 If taskFinishIdx > levelMaxFinish(currentLevel) Then
                    levelMaxFinish(currentLevel) = taskFinishIdx
                    levelMaxFinishAlloc(currentLevel) = taskFinishAlloc
                ElseIf taskFinishIdx = levelMaxFinish(currentLevel) Then
                    If taskFinishAlloc > levelMaxFinishAlloc(currentLevel) Then
                        levelMaxFinishAlloc(currentLevel) = taskFinishAlloc
                    End If
                End If
            End If
            
        Else
            ' Unlocked -> Schedule it
            duration = 0
            If IsNumeric(metaData(taskRow, COL_DURATION_IDX)) Then duration = CDbl(metaData(taskRow, COL_DURATION_IDX))
            
            ' Clear grid for this row
            For dayIdx = 1 To numDays
                gridData(taskRow, dayIdx) = Empty
            Next dayIdx
            
            If assigneeName <> "" And duration > 0 Then
                If Not personUsage.Exists(assigneeName) Then
                    ReDim newAllocArray(1 To numDays) As Double
                    personUsage.Add assigneeName, newAllocArray
                End If
                
                ' Get Max Daily Load for Person
                maxDailyLoad = 1#
                If capacityLimits.Exists(assigneeName) Then
                    maxDailyLoad = capacityLimits(assigneeName)
                End If
                
                remaining = duration
                newAllocArray = personUsage(assigneeName)
                
                taskStartIdx = baseStartIdx
                
                ' Add Lag (Start Offset) from Column E
                lagDays = metaData(taskRow, COL_OFFSET_IDX)
                If IsNumeric(lagDays) And Not IsEmpty(lagDays) Then
                    taskStartIdx = taskStartIdx + CLng(lagDays)
                End If
                
                If taskStartIdx < 1 Then taskStartIdx = 1
                
                ' Check if Micro-Task (Standard rounding to 0 but original > 0)
                Dim totalNeeded As Long
                Dim isMicroTask As Boolean
                totalNeeded = Int((duration / 0.25) + 0.5)
                isMicroTask = (totalNeeded = 0 And duration > 0)
                
                ' Allocate loop
                For dayIdx = taskStartIdx To numDays
                    If remaining <= 0 Then Exit For
                    
                    isHoliday = (Trim(holidayData(1, dayIdx)) = STR_HOLIDAY_MARK)
                    
                    If Not isHoliday Then
                        ' Capacity based on Configured Limit
                        capacity = maxDailyLoad - newAllocArray(dayIdx)
                        If capacity < 0 Then capacity = 0
                        
                        Dim maxUnits As Long
                        maxUnits = Int(capacity / 0.25)
                        
                        Dim neededUnits As Long
                        Dim allocateUnits As Long
                        dailyAlloc = 0
                        
                        If isMicroTask Then
                             ' Special Case: Micro-Task (e.g. 0.1, 0.05)
                             ' Only allocate if capacity >= 0.1
                             ' If allocated, it consumes 0.1 and we are done.
                             If capacity >= 0.1 Then
                                dailyAlloc = 0.1
                             End If
                        Else
                             ' Standard Logic (0.25 units)
                             neededUnits = Int((remaining / 0.25) + 0.5)
                             
                             If maxUnits > 0 And neededUnits > 0 Then
                                allocateUnits = neededUnits
                                If allocateUnits > maxUnits Then allocateUnits = maxUnits
                                dailyAlloc = allocateUnits * 0.25
                             End If
                        End If
                            
                        If dailyAlloc > 0 Then
                            gridData(taskRow, dayIdx) = dailyAlloc
                            newAllocArray(dayIdx) = newAllocArray(dayIdx) + dailyAlloc
                            remaining = remaining - dailyAlloc
                            
                            ' Update row finish tracking
                            taskFinishIdx = dayIdx
                            taskFinishAlloc = dailyAlloc
                        End If
                    End If
                Next dayIdx
                
                personUsage(assigneeName) = newAllocArray
            End If
            
            ' Update Level Max Finish Logic for Unlocked Row
            If currentLevel > 0 Then
                If taskFinishIdx > levelMaxFinish(currentLevel) Then
                    levelMaxFinish(currentLevel) = taskFinishIdx
                    levelMaxFinishAlloc(currentLevel) = taskFinishAlloc
                ElseIf taskFinishIdx = levelMaxFinish(currentLevel) Then
                    If taskFinishAlloc > levelMaxFinishAlloc(currentLevel) Then
                        levelMaxFinishAlloc(currentLevel) = taskFinishAlloc
                    End If
                End If
            End If
        End If
        
    Next taskRow
    
    ' 4. Write Back to Sheet
    rangeGrid.Value = gridData
    
    ' MsgBox "Scheduling Complete!"
    
Cleanup:
    Application.ScreenUpdating = screenUpdateState
    Application.Calculation = calcState
    Application.EnableEvents = eventsState
    
End Sub
