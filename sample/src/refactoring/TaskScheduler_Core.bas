Option Explicit

' カレンダーの位置と属性
Type CalendarConfig
    ROW_HEADER As Long
    ROW_HOLIDAY As Long
    COL_CALENDAR_START As Long
    STR_HOLIDAY_MARK As String
End Type

' タスクの位置と属性
Type TaskConfig
    ROW_START As Long
    COL_LEVEL As Long
    COL_OFFSET As Long
    COL_LOCK As Long
    COL_DURATION As Long
    COL_ASSIGNEE As Long
    STR_LOCK_MARK As String
End Type

' 担当者の位置と属性
Type AssigneeConfig
    COL_NAME As Long
    COL_LIMIT As Long
    ROW_START As Long
    ROW_END As Long
End Type

Function InitCalendarConfig() As CalendarConfig
    Dim cfg As CalendarConfig
    cfg.ROW_HEADER = 3
    cfg.ROW_HOLIDAY = 5
    cfg.COL_CALENDAR_START = 24
    cfg.STR_HOLIDAY_MARK = "休"
    InitCalendarConfig = cfg
End Function

Function InitTaskConfig() As TaskConfig
    Dim cfg As TaskConfig
    cfg.ROW_START = 19
    cfg.COL_LEVEL = 8
    cfg.COL_OFFSET = 9
    cfg.COL_LOCK = 13
    cfg.COL_DURATION = 15
    cfg.COL_ASSIGNEE = 17
    cfg.COL_END = 17
    cfg.STR_LOCK_MARK = "L"
    InitTaskConfig = cfg
End Function

Function InitAssigneeConfig() As AssigneeConfig
    Dim cfg As AssigneeConfig
    cfg.COL_NAME = 17
    cfg.COL_LIMIT = 18
    cfg.ROW_START = 8
    cfg.ROW_END = 14
    InitAssigneeConfig = cfg
End Function

' =========================================================
' Getter Functions: Configフィールドのカプセル化
' =========================================================

' --- Boundary Detection (境界検出) ---

Function GetLastTaskRow(ws As Worksheet, taskCfg As TaskConfig) As Long
    GetLastTaskRow = ws.Cells(ws.Rows.Count, taskCfg.COL_DURATION).End(xlUp).Row
End Function

Function GetLastCalendarCol(ws As Worksheet, calCfg As CalendarConfig) As Long
    GetLastCalendarCol = ws.Cells(calCfg.ROW_HEADER, ws.Columns.Count).End(xlToLeft).Column
End Function

Function CalcNumRows(lastRow As Long, taskCfg As TaskConfig) As Long
    CalcNumRows = lastRow - taskCfg.ROW_START + 1
End Function

Function CalcNumDays(lastCol As Long, calCfg As CalendarConfig) As Long
    CalcNumDays = lastCol - calCfg.COL_CALENDAR_START + 1
End Function

' --- Range Building (Range構築) ---

Function GetTaskRange(ws As Worksheet, taskCfg As TaskConfig, lastRow As Long) As Range
    Set GetTaskRange = ws.Range(ws.Cells(taskCfg.ROW_START, 1), ws.Cells(lastRow, taskCfg.COL_END))
End Function

Function GetScheduleRange(ws As Worksheet, taskCfg As TaskConfig, calCfg As CalendarConfig, lastRow As Long, lastCol As Long) As Range
    Set GetScheduleRange = ws.Range(ws.Cells(taskCfg.ROW_START, calCfg.COL_CALENDAR_START), ws.Cells(lastRow, lastCol))
End Function

Function GetHolidayRange(ws As Worksheet, calCfg As CalendarConfig, lastCol As Long) As Range
    Set GetHolidayRange = ws.Range(ws.Cells(calCfg.ROW_HOLIDAY, calCfg.COL_CALENDAR_START), ws.Cells(calCfg.ROW_HOLIDAY, lastCol))
End Function

Function GetAssigneeRange(ws As Worksheet, assigneeCfg As AssigneeConfig) As Range
    Set GetAssigneeRange = ws.Range(ws.Cells(assigneeCfg.ROW_START, assigneeCfg.COL_NAME), ws.Cells(assigneeCfg.ROW_END, assigneeCfg.COL_LIMIT))
End Function

' --- Per-Row Field Reads (行フィールド読取) ---

Function IsRowLocked(taskDataFrame As Variant, taskRow As Long, taskCfg As TaskConfig) As Boolean
    IsRowLocked = (UCase(Trim(taskDataFrame(taskRow, taskCfg.COL_LOCK))) = taskCfg.STR_LOCK_MARK)
End Function

Function GetAssigneeName(taskDataFrame As Variant, taskRow As Long, taskCfg As TaskConfig) As String
    GetAssigneeName = Trim(taskDataFrame(taskRow, taskCfg.COL_ASSIGNEE))
End Function

' グリッドセルの値を Double として返す。空・非数値の場合は 0 を返す。
Function GetNumericCellValue(grid As Variant, r As Long, c As Long) As Double
    Dim v As Variant
    v = grid(r, c)
    If IsNumeric(v) And Not IsEmpty(v) Then GetNumericCellValue = CDbl(v) Else GetNumericCellValue = 0
End Function

Function GetTaskLevel(taskDataFrame As Variant, taskRow As Long, taskCfg As TaskConfig) As Long
    GetTaskLevel = 0
    If IsNumeric(taskDataFrame(taskRow, taskCfg.COL_LEVEL)) And Not IsEmpty(taskDataFrame(taskRow, taskCfg.COL_LEVEL)) Then
        GetTaskLevel = CLng(taskDataFrame(taskRow, taskCfg.COL_LEVEL))
    End If
End Function

Function GetMaxLevel(taskDataFrame As Variant, numRows As Long, taskCfg As TaskConfig) As Long
    Dim maxLvl As Long
    Dim lvl As Long
    Dim taskRow As Long
    maxLvl = 0
    For taskRow = 1 To numRows
        lvl = GetTaskLevel(taskDataFrame, taskRow, taskCfg)
        If lvl > maxLvl Then maxLvl = lvl
    Next taskRow
    GetMaxLevel = maxLvl
End Function

' Refactor #1: Extract Base Start Index Calculation
' 依存タスクの開始日を計算する (Calculate the start date for dependent tasks)
' Logic: If parentFinishAlloc < 0.5, start on parentFinishIdx. Else, parentFinishIdx + 1.
' ロジック: 親タスクの最終日の割り当てが0.5未満なら親の最終日と同日に開始、それ以外は翌日から開始する
Function CalcBaseStartIdx(currentLevel As Long, parentFinishIdx As Long, parentFinishAlloc As Double) As Long
    Dim baseStartIdx As Long
    baseStartIdx = 1

    If currentLevel > 1 Then
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
    
    CalcBaseStartIdx = baseStartIdx
End Function

' Refactor #2: Extract Max Daily Load Lookup
' 担当者の1日あたりの最大稼働工数を取得する (Get the maximum daily workload for an assignee)
' Logic: Look up assignee in limits dict, default to 1.0.
' ロジック: キャパシティ設定辞書から担当者を検索し、存在しなければデフォルト値 1.0 を返す
Function GetMaxDailyLoad(assigneeName As String, capacityLimits As Object) As Double
    Dim maxDailyLoad As Double
    maxDailyLoad = 1# ' Default 1.0

    If capacityLimits.Exists(assigneeName) Then
        maxDailyLoad = capacityLimits(assigneeName)
    End If

    GetMaxDailyLoad = maxDailyLoad
End Function

' Refactor #3: Extract Daily Allocation Calculation
' 1日あたりのタスク割り当て工数を計算する (Calculate daily task allocation)
' Logic: Handle 0.25 unit allocations and the 0.1 micro-task minimum based on capacity.
' ロジック: 残り容量と工数から0.25単位の標準割り当てを計算する。マイクロタスクの場合は最低0.1を保証する
Function CalcDailyAllocation(capacity As Double, remaining As Double, isMicroTask As Boolean) As Double
    Dim dailyAllocation As Double
    dailyAllocation = 0
    
    Dim maxUnits As Long
    maxUnits = Int(capacity / 0.25)
    
    Dim neededUnits As Long
    Dim allocateUnits As Long
    
    If isMicroTask Then
         ' Special Case: Micro-Task (e.g. 0.1, 0.05)
         ' Only allocate if capacity >= 0.1
         ' If allocated, it consumes 0.1 and we are done.
         If capacity >= 0.1 Then
            dailyAllocation = 0.1
         End If
    Else
         ' Standard Logic (0.25 units)
         neededUnits = Int((remaining / 0.25) + 0.5)
         
         If maxUnits > 0 And neededUnits > 0 Then
            allocateUnits = neededUnits
            If allocateUnits > maxUnits Then allocateUnits = maxUnits
            dailyAllocation = allocateUnits * 0.25
         End If
    End If
    
    CalcDailyAllocation = dailyAllocation
End Function

' Refactor #4: Extract Level Finish Update Logic
' 指定された階層レベルの最終完了日（最も遅い日）とその日の工数を更新する (Update the maximum finish date and its allocation for a given level)
' Logic: Update the max finish index and allocation arrays based on the task's completion.
' ロジック: タスクの完了日が現在の最大値より後ろなら更新、同日なら工数が多い方で上書きする
Sub UpdateLevelFinish(currentLevel As Long, taskFinishIdx As Long, taskFinishAlloc As Double, ByRef levelMaxFinish, ByRef levelMaxFinishAlloc)
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
End Sub

' Refactor #5: Extract Capacity Config Building
' 設定データ（2D配列）から、担当者名と1日の最大稼働工数（デフォルト1.0）を紐づけたディクショナリを生成する
' Logic: Parse a 2D array [Name, Capacity] and return a Scripting.Dictionary 
Function BuildCapacityDict(assigneeDataFrame As Variant) As Object
    Dim capacityLimits As Object
    Set capacityLimits = CreateObject("Scripting.Dictionary")
    
    Dim cfgRow As Long
    Dim cfgCapacity As Double
    Dim cfgName As String
    
    For cfgRow = 1 To UBound(assigneeDataFrame, 1)
        cfgName = Trim(assigneeDataFrame(cfgRow, 1))
        If cfgName <> "" Then
            cfgCapacity = 1# ' Default
            If IsNumeric(assigneeDataFrame(cfgRow, 2)) And Not IsEmpty(assigneeDataFrame(cfgRow, 2)) Then
                cfgCapacity = CDbl(assigneeDataFrame(cfgRow, 2))
            End If
            capacityLimits(cfgName) = cfgCapacity
        End If
    Next cfgRow
    
    Set BuildCapacityDict = capacityLimits
End Function

' Refactor #6: Extract Phase 1 (Scan Locked Rows)
' 全タスクをスキャンし、「L」マーク（ロック）がついている場合は既存スケジュールを assigneeUsage (実績Dict) に事前割り当てする
' ロック行の日別グリッド値を newAllocArray に加算する
Sub AccumulateLockedRowUsage(taskRow As Long, numDays As Long, scheduleGrid As Variant, ByRef newAllocArray() As Double)
    Dim dayIdx As Long
    Dim existingAlloc As Double
    For dayIdx = 1 To numDays
        existingAlloc = GetNumericCellValue(scheduleGrid, taskRow, dayIdx)
        If existingAlloc > 0 Then
            newAllocArray(dayIdx) = newAllocArray(dayIdx) + existingAlloc
        End If
    Next dayIdx
End Sub

Sub ScanLockedRows(taskCfg As TaskConfig, numRows As Long, numDays As Long, taskDataFrame As Variant, scheduleGrid As Variant, ByRef assigneeUsage As Object)
    Dim taskRow As Long
    Dim assigneeName As String
    Dim newAllocArray() As Double

    For taskRow = 1 To numRows
        assigneeName = Trim(taskDataFrame(taskRow, taskCfg.COL_ASSIGNEE))

        If assigneeName <> "" Then
            If Not assigneeUsage.Exists(assigneeName) Then
                ReDim newAllocArray(1 To numDays) As Double
                assigneeUsage.Add assigneeName, newAllocArray
            End If

            If UCase(Trim(taskDataFrame(taskRow, taskCfg.COL_LOCK))) = taskCfg.STR_LOCK_MARK Then
                newAllocArray = assigneeUsage(assigneeName)
                Call AccumulateLockedRowUsage(taskRow, numDays, scheduleGrid, newAllocArray)
                assigneeUsage(assigneeName) = newAllocArray
            End If
        End If
    Next taskRow
End Sub

' Refactor #7: Extract ClearTaskGridRow
' 指定されたタスク行のスケジュールグリッド（右側）をクリアする
Sub ClearTaskGridRow(taskRow As Long, numDays As Long, ByRef scheduleGrid As Variant)
    Dim dayIdx As Long
    For dayIdx = 1 To numDays
        scheduleGrid(taskRow, dayIdx) = Empty
    Next dayIdx
End Sub

' Refactor #8: Extract FindLockedTaskFinish
' ロックされたタスク行を右から左へスキャンし、最後に割り当てがある日 (taskFinishIdx) とその量 (taskFinishAlloc) を見つける
Sub FindLockedTaskFinish(ByVal taskRow As Long, ByVal numDays As Long, ByRef scheduleGrid As Variant, ByRef taskFinishIdx As Long, ByRef taskFinishAlloc As Double)
    Dim dayIdx As Long
    Dim existingAlloc As Double

    taskFinishIdx = 0
    taskFinishAlloc = 0

    For dayIdx = numDays To 1 Step -1
        existingAlloc = GetNumericCellValue(scheduleGrid, taskRow, dayIdx)
        If existingAlloc > 0 Then
            taskFinishIdx = dayIdx
            taskFinishAlloc = existingAlloc
            Exit For
        End If
    Next dayIdx
End Sub

' Refactor #9: Extract ScheduleUnlockedTask
' 個別のアンロック済みタスクスケジュール（日別の工数割り当てなど）処理を実行する
Sub ScheduleUnlockedTask(taskCfg As TaskConfig, calCfg As CalendarConfig, ByVal taskRow As Long, ByVal numDays As Long, ByVal baseStartIdx As Long, ByVal taskDataFrame As Variant, ByVal holidayData As Variant, ByVal capacityLimits As Object, ByRef scheduleGrid As Variant, ByRef assigneeUsage As Object, ByRef taskFinishIdx As Long, ByRef taskFinishAlloc As Double)
    Dim duration As Double
    Dim assigneeName As String
    Dim remaining As Double
    Dim dailyAllocation As Double
    Dim capacity As Double
    Dim maxDailyLoad As Double
    Dim isHoliday As Boolean
    
    Dim lagDays As Variant
    Dim taskStartIdx As Long
    Dim dayIdx As Long
    
    Dim newAllocArray() As Double
    
    assigneeName = Trim(taskDataFrame(taskRow, taskCfg.COL_ASSIGNEE))
    duration = 0
    If IsNumeric(taskDataFrame(taskRow, taskCfg.COL_DURATION)) Then duration = CDbl(taskDataFrame(taskRow, taskCfg.COL_DURATION))
    
    ' Clear grid for this row
    Call ClearTaskGridRow(taskRow, numDays, scheduleGrid)
    
    If assigneeName <> "" And duration > 0 Then
        If Not assigneeUsage.Exists(assigneeName) Then
            ReDim newAllocArray(1 To numDays) As Double
            assigneeUsage.Add assigneeName, newAllocArray
        End If
        
        ' Get Max Daily Load for Person
        maxDailyLoad = GetMaxDailyLoad(assigneeName, capacityLimits)
        
        remaining = duration
        newAllocArray = assigneeUsage(assigneeName)
        
        taskStartIdx = baseStartIdx
        
        ' Add Lag (Start Offset)
        lagDays = taskDataFrame(taskRow, taskCfg.COL_OFFSET)
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
            
            isHoliday = (Trim(holidayData(1, dayIdx)) = calCfg.STR_HOLIDAY_MARK)
            
            If Not isHoliday Then
                ' Capacity based on Configured Limit
                capacity = maxDailyLoad - newAllocArray(dayIdx)
                If capacity < 0 Then capacity = 0
                
                dailyAllocation = CalcDailyAllocation(capacity, remaining, isMicroTask)
                    
                If dailyAllocation > 0 Then
                    scheduleGrid(taskRow, dayIdx) = dailyAllocation
                    newAllocArray(dayIdx) = newAllocArray(dayIdx) + dailyAllocation
                    remaining = remaining - dailyAllocation
                    
                    ' Update row finish tracking
                    taskFinishIdx = dayIdx
                    taskFinishAlloc = dailyAllocation
                End If
            End If
        Next dayIdx
        
        assigneeUsage(assigneeName) = newAllocArray
    End If
End Sub

' ---------------------------------------------------------
' Test Wrappers (Used by TaskScheduler_Core.test.ts)
' ---------------------------------------------------------

Function TestFindLockedTaskFinish(taskRow As Long, numDays As Long, ByRef scheduleGrid As Variant) As String
    Dim fIdx As Long
    Dim fAlloc As Double
    Call FindLockedTaskFinish(taskRow, numDays, scheduleGrid, fIdx, fAlloc)
    TestFindLockedTaskFinish = fIdx & "|" & fAlloc
End Function

Function TestScheduleUnlockedTask(taskCfg As TaskConfig, calCfg As CalendarConfig, taskRow As Long, numDays As Long, baseStartIdx As Long, taskDataFrame As Variant, holidayData As Variant, capacityLimits As Object, ByRef scheduleGrid As Variant, ByRef assigneeUsage As Object) As String
    Dim fIdx As Long
    Dim fAlloc As Double
    Call ScheduleUnlockedTask(taskCfg, calCfg, taskRow, numDays, baseStartIdx, taskDataFrame, holidayData, capacityLimits, scheduleGrid, assigneeUsage, fIdx, fAlloc)
    TestScheduleUnlockedTask = fIdx & "|" & fAlloc
End Function
