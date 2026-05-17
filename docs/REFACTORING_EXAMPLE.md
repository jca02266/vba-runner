# VBA リファクタリング具体例：TaskScheduler マクロ

## 概要

このドキュメントは、**実際のレガシーコード（TaskScheduler_v1.bas）をリファクタリングして（TaskScheduler.bas + TaskScheduler_Core.bas）どのように改善されたか、具体的に示します**。

### 対象コード

- **リファクタリング前**: `sample/src/vba_legacy/TaskScheduler_v1.bas`（393 行の巨大 Sub）
- **リファクタリング後**: `sample/src/vba/TaskScheduler.bas` + `sample/src/vba/TaskScheduler_Core.bas`

`TaskScheduler_v1.bas` が何をするマクロなのか（シートレイアウト・使い方・動作ルール）は [`sample/src/vba_legacy/TaskScheduler.md`](../sample/src/vba_legacy/TaskScheduler.md) を参照してください。

---

## 問題点：リファクタリング前のコード

### 問題 1: 単一の巨大 Sub（393 行）

```vba
' TaskScheduler_v1.bas
Sub AutoScheduleTasks()
    ' Const 定義 (50 行)
    Const COL_CALENDAR_START_IDX As Long = 24
    Const ROW_HEADER_SHEET As Long = 3
    ' ... 15+ の定数定義
    
    ' 変数宣言 (30 行)
    Dim ws As Worksheet
    Dim screenUpdateState As Boolean
    ' ... 複数の変数
    
    ' Phase 1: Scan Locked Rows （30 行）
    For taskRow = 1 To numRows
        ' ネストされた If/For ループ
        ' 状態変数を更新しながら処理
    Next
    
    ' Phase 2: Schedule & Calculate Dependencies (280 行)
    For taskRow = 1 To numRows
        ' 階層ロジック
        ' 依存関係計算
        ' スケジューリング
        ' Level Max 更新
        ' ... すべてが混在
    Next
    
    ' Phase 3: Write Back
    rangeGrid.Value = gridData
End Sub
```

### 問題 2: 読み方が一方向（上から下へ）

レガシーコードを理解するには、**すべての行を順番に読む必要がある**：

```
1行目 ~ 50行目: Const 定義を暗記
51行目 ~ 85行目: 変数宣言を追跡
86行目: Range 計算開始
100行目: あ、ここで配列初期化か
...
250行目: あ、ここで Locked Row の処理だ
280行目: ここから本メインロジック？
...
```

**結果**: 
- 全 393 行を読んでようやく全体像がわかる
- 「ここは何をしているのか」が不明確
- 変更時に影響範囲が不明（どこが依存しているのか？）
- **デバッグが困難**（どの部分で状態が狂ったのか？）

### 問題 3: 状態管理の混在

```vba
' Phase 1 と Phase 2 で状態が共有される
Dim levelMaxFinish(0 To 100) As Long        ' Level ごとの最終日
Dim personUsage As Object                   ' 担当者ごとの工数使用量
Dim gridData As Variant                     ' スケジュール配列

' Phase 1 で以下を更新
For taskRow = 1 To numRows
    ' personUsage を更新
    personUsage(assigneeName) = newAllocArray
Next

' Phase 2 で上記の personUsage に依存しながら処理
For taskRow = 1 To numRows
    ' personUsage を読み取りながら
    capacity = maxDailyLoad - newAllocArray(dayIdx)
    ' ...同時に gridData も更新
Next
```

**問題**: 
- 複数の Variant 配列と Dictionary が共有される
- どこで何が変更されるのか追跡困難
- バグの原因を特定しづらい

---

## TaskScheduler_v1.bas の処理概要

このセクションでは、リファクタリング前の TaskScheduler_v1.bas（393 行の巨大 Sub）がどのような処理フローを実装しているのか、詳細に説明します。このコードは実業務で必要な複雑なロジックを含んでいますが、すべてが 1 つの Sub に詰め込まれているため、その複雑さが明らかになりにくくなっています。

### フェーズ 0: 初期化と最適化（行 49-84）

**定数定義（50 個以上）**:
```vba
Const COL_CALENDAR_START_IDX As Long = 24  ' スケジュール開始列
Const ROW_HEADER_SHEET As Long = 3         ' ヘッダー行
Const ROW_HOLIDAY_SHEET As Long = 5        ' 休日マーク行
Const COL_LEVEL_IDX As Long = 8            ' 階層（Level）列
Const COL_OFFSET_IDX As Long = 9           ' 開始遅延日数（Lag）列
Const COL_LOCK_IDX As Long = 13            ' ロック状態列
Const COL_DURATION_IDX As Long = 15        ' 工数（Duration）列
Const COL_ASSIGNEE_IDX As Long = 17        ' 担当者名列
' ... 40+ の定数が scattered
```

**パフォーマンス最適化**:
```vba
screenUpdateState = Application.ScreenUpdating
calcState = Application.Calculation
eventsState = Application.EnableEvents

Application.ScreenUpdating = False          ' 画面更新を無効化
Application.Calculation = xlCalculationManual  ' 自動計算を無効化
Application.EnableEvents = False            ' イベント処理を無効化
```

### フェーズ 1: 範囲検出（行 87-99）

Excel シートの**実際のデータ範囲**を検出：
```vba
lastRow = ws.Cells(ws.Rows.Count, COL_DURATION_IDX).End(xlUp).Row
lastCol = ws.Cells(ROW_HEADER_SHEET, ws.Columns.Count).End(xlToLeft).Column
numRows = lastRow - ROW_START_SHEET + 1
numDays = lastCol - COL_CALENDAR_START_IDX + 1
```

**目的**: タスク数とカレンダー日数を動的に計算

### フェーズ 2: データ読み込み（行 101-122） [[→ R-05](REFACTORING_TESTING_CATALOG.md#r-05)]

**Double Buffering パターン** — Excel シートをメモリの配列に読み込む：

```vba
' メタデータ（タスク情報）
Set rangeMeta = ws.Range(ws.Cells(ROW_START_SHEET, 1), _
                         ws.Cells(lastRow, CONFIG_COL_NAME))
metaData = rangeMeta.Value  ' 2D配列に展開

' スケジュール結果格納先
Set rangeGrid = ws.Range(ws.Cells(ROW_START_SHEET, COL_CALENDAR_START_IDX), _
                         ws.Cells(lastRow, lastCol))
gridData = rangeGrid.Value  ' 2D配列に展開

' 休日マーク
Set rangeHoliday = ws.Range(ws.Cells(ROW_HOLIDAY_SHEET, COL_CALENDAR_START_IDX), _
                            ws.Cells(ROW_HOLIDAY_SHEET, lastCol))
holidayData = rangeHoliday.Value
```

**重要**: すべての処理がメモリ上の配列で行われ、最後に一度だけシートに書き戻される（パフォーマンス最適化）

### フェーズ 3: キャパシティ設定の読み込み（行 117-139）

各担当者の**最大工数（キャパシティ）** をセルから読み込み、Dictionary に格納：

```vba
Set capacityLimits = CreateObject("Scripting.Dictionary")

For cfgRow = 1 To UBound(configData, 1)
    cfgName = Trim(configData(cfgRow, 1))  ' 担当者名
    cfgCapacity = 1.0  ' デフォルト
    If IsNumeric(configData(cfgRow, 2)) Then
        cfgCapacity = CDbl(configData(cfgRow, 2))  ' 設定値で上書き
    End If
    capacityLimits(cfgName) = cfgCapacity
Next cfgRow
```

**結果**: `capacityLimits("Alice") = 1.0, capacityLimits("Bob") = 0.8` のような Dictionary

### フェーズ 4: Phase 1 — Locked Rows スキャン（行 151-179）

**ロック行（"L" マーク）** の工数を集計。これらはスケジュール変更の対象外だが、**リソース占有を計算する**ために追跡される：

```vba
Dim personUsage As Object
Set personUsage = CreateObject("Scripting.Dictionary")

For taskRow = 1 To numRows
    assigneeName = Trim(metaData(taskRow, COL_ASSIGNEE_IDX))
    
    If assigneeName <> "" Then
        ' 初回出現時に配列を初期化
        If Not personUsage.Exists(assigneeName) Then
            ReDim newAllocArray(1 To numDays) As Double
            personUsage.Add assigneeName, newAllocArray
        End If
        
        ' Locked Row なら工数を集計
        If UCase(Trim(metaData(taskRow, COL_LOCK_IDX))) = "L" Then
            For dayIdx = 1 To numDays
                existingAlloc = CDbl(gridData(taskRow, dayIdx))
                newAllocArray(dayIdx) = newAllocArray(dayIdx) + existingAlloc
            Next dayIdx
            personUsage(assigneeName) = newAllocArray
        End If
    End If
Next taskRow
```

**結果**: `personUsage("Alice") = [0.5, 0.5, 1.0, 0.5, ...]` という工数配列

### フェーズ 5: Phase 2 — スケジューリング（行 181-381）

最も複雑なフェーズ。各タスク行について以下を実行：

#### 5a. Level 依存関係の計算（行 211-241）

```vba
Dim levelMaxFinish(0 To 100) As Long        ' Level ごとの完了日
Dim levelMaxFinishAlloc(0 To 100) As Double ' 完了日の割当量

currentLevel = CLng(metaData(taskRow, COL_LEVEL_IDX))

' Level 1 -> 新しいタスクブロック開始（依存関係をリセット）
If currentLevel = 1 Then
    Erase levelMaxFinish
    Erase levelMaxFinishAlloc
End If

' Level > 1 -> 親タスクの完了状況に基づいて開始日を計算
baseStartIdx = 1
If currentLevel > 1 Then
    parentFinishIdx = levelMaxFinish(currentLevel - 1)
    parentFinishAlloc = levelMaxFinishAlloc(currentLevel - 1)
    
    ' 重要な仕様: 親の完了日の割当が 0.5 未満なら「同日開始」
    If parentFinishAlloc < 0.5 Then
        baseStartIdx = parentFinishIdx
    Else
        baseStartIdx = parentFinishIdx + 1  ' それ以外は「翌日開始」
    End If
End If
```

**複雑な状態管理**: `levelMaxFinish` 配列を追跡しながら、すべてのタスク行を一度に処理

#### 5b. Locked タスク（行 249-274）

Locked 行は**スケジュール変更対象外**だが、**Level Max を更新**する必要がある：

```vba
If isLocked Then
    ' 右から左へスキャン（最後の割当を見つける）
    For dayIdx = numDays To 1 Step -1
        If CDbl(gridData(taskRow, dayIdx)) > 0 Then
            taskFinishIdx = dayIdx
            taskFinishAlloc = CDbl(gridData(taskRow, dayIdx))
            Exit For
        End If
    Next dayIdx
    
    ' Level Max を更新
    If taskFinishIdx > levelMaxFinish(currentLevel) Then
        levelMaxFinish(currentLevel) = taskFinishIdx
        levelMaxFinishAlloc(currentLevel) = taskFinishAlloc
    End If
End If
```

#### 5c. Unlocked タスク — スケジューリング（行 276-379）

スケジュール対象のタスク。最も複雑なロジック：

**開始日の計算**:
```vba
taskStartIdx = baseStartIdx
' Start Offset（Lag 日数）を追加
lagDays = metaData(taskRow, COL_OFFSET_IDX)
If IsNumeric(lagDays) Then
    taskStartIdx = taskStartIdx + CLng(lagDays)
End If
If taskStartIdx < 1 Then taskStartIdx = 1
```

**マイクロタスク検出** — 工数が 0.25 単位で丸めると 0 になる場合（例: 0.1, 0.05）:
```vba
Dim totalNeeded As Long
totalNeeded = Int((duration / 0.25) + 0.5)
isMicroTask = (totalNeeded = 0 And duration > 0)
```

**割当ループ** — 開始日から最終日まで工数を割り当て：
```vba
remaining = duration
For dayIdx = taskStartIdx To numDays
    If remaining <= 0 Then Exit For
    
    ' 休日チェック
    isHoliday = (Trim(holidayData(1, dayIdx)) = "休")
    If isHoliday Then GoTo NextDay
    
    ' 担当者のキャパシティを計算
    capacity = maxDailyLoad - newAllocArray(dayIdx)
    
    If isMicroTask Then
        ' 特殊: 0.1 以上の容量があれば 0.1 を割当
        If capacity >= 0.1 Then
            dailyAlloc = 0.1
        End If
    Else
        ' 通常: 0.25 単位で割当
        neededUnits = Int((remaining / 0.25) + 0.5)
        maxUnits = Int(capacity / 0.25)
        allocateUnits = Min(neededUnits, maxUnits)
        dailyAlloc = allocateUnits * 0.25
    End If
    
    If dailyAlloc > 0 Then
        gridData(taskRow, dayIdx) = dailyAlloc
        newAllocArray(dayIdx) = newAllocArray(dayIdx) + dailyAlloc
        remaining = remaining - dailyAlloc
        taskFinishIdx = dayIdx
        taskFinishAlloc = dailyAlloc
    End If
NextDay:
Next dayIdx

personUsage(assigneeName) = newAllocArray
```

**Level Max 更新**:
```vba
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
```

### フェーズ 6: 書き戻し（行 384）

メモリ上の配列をシートに一度だけ書き戻し：
```vba
rangeGrid.Value = gridData
```

### フェーズ 7: クリーンアップ（行 388-391）

Application 設定を復元：
```vba
Application.ScreenUpdating = screenUpdateState
Application.Calculation = calcState
Application.EnableEvents = eventsState
```

### 複雑度の総括

| 項目 | 説明 |
|------|------|
| 定数 | 50+ 個がバラバラに定義 |
| 状態変数 | metaData, gridData, personUsage, levelMaxFinish, levelMaxFinishAlloc など |
| ネスト深さ | 最大 4 層（task loop > locked check > day loop > unit loop） |
| 条件分岐 | Level 判定、Locked/Unlocked 分岐、マイクロタスク分岐、休日判定 |
| 配列操作 | 2D配列（metaData, gridData, configData, holidayData） + Dictionary（personUsage, capacityLimits） |
| 総行数 | 393 行が 1 つの Sub 内に収納 |

**読者が保持すべき「心的モデル」**:
- 50 個の定数の役割
- 6 つの配列/Dictionary の相互関係
- Phase 1 と Phase 2 の依存関係
- Level 計算の仕様（0.5 閾値）
- マイクロタスクの特殊処理
- 0.25 単位の容量管理
- 状態変数（levelMaxFinish）の更新タイミング

すべてを理解するには **全 393 行を順に読む必要がある** というのが、リファクタリング前の最大の問題です。

### Excel との密結合によるテスト困難性

TaskScheduler_v1.bas の最も深刻な問題が、**Excel との密結合**です。ビジネスロジックと I/O 操作が混在しているため、**VBA Runner で単体テスト不可能**になっています。

#### 問題 1: Excel シート依存

```vba
' フェーズ 2 の開始時
Set rangeMeta = ws.Range(ws.Cells(ROW_START_SHEET, 1), _
                         ws.Cells(lastRow, CONFIG_COL_NAME))
metaData = rangeMeta.Value

' フェーズ 5 の割当ループ内で常に Excel に参照
isHoliday = (Trim(holidayData(1, dayIdx)) = STR_HOLIDAY_MARK)
' ...

' フェーズ 6 の書き戻し
rangeGrid.Value = gridData
```

**テストの困難性**:
- `ws.Range()` を呼び出すには **実際の Excel ファイルが必要**
- VBA 実行エンジン（例: Node.js の VBA Runner）で実行すると、`Range` オブジェクトが undefined
- Mock なしでは、**Excel を起動してファイルを手動で開く** という手段しかない

#### 問題 2: ビジネスロジック × I/O の密結合

```vba
' レベル計算（ビジネスロジック）が Excel 読み込み（I/O）と混在
For taskRow = 1 To numRows  ' ← numRows は Excel から読んだ値
    assigneeName = Trim(metaData(taskRow, COL_ASSIGNEE_IDX))  ' ← metaData は Range.Value
    currentLevel = CLng(metaData(taskRow, COL_LEVEL_IDX))     ' ← 同じ配列から読む
    
    ' ここからビジネスロジック（純粋な計算）が始まるが...
    If currentLevel > 1 Then
        baseStartIdx = levelMaxFinish(currentLevel - 1)  ' ビジネスロジック
        ' ...複雑な計算...
        
        ' 途中で I/O 操作が混在
        capacity = maxDailyLoad - newAllocArray(dayIdx)
        gridData(taskRow, dayIdx) = dailyAlloc  ' ← 配列への書き込み
    End If
Next taskRow

' 最後に全体を一度に書き戻す
rangeGrid.Value = gridData  ' ← I/O 操作
```

**問題**:
- **テスト時には I/O をスキップしたい** が、ロジック内に深く埋まっている
- 「レベル計算だけをテストしたい」と思っても、全 393 行が連動している
- 配列データの準備と検証が複雑（Excel が必須）

#### 問題 3: 定義上のテスト不可能性

```typescript
// こういう単体テストが書きたい
function testCalcBaseStartIdx() {
  // 「親タスクが 3 日目、割当 0.3 なら、同日開始」
  const result = CalcBaseStartIdx(2, 3, 0.3);
  assert.equal(result, 3);  // 期待値: 3 日目
}
```

**しかし VBA ではこのテストが書けない**:
- `CalcBaseStartIdx` が Sub 内にネストされているので抽出不可
- `metaData`, `gridData` 等の配列が Sub スコープのため外部アクセス不可
- 関数を分離しても、Excel Range への参照が残っているテストは Excel なしで実行不可

#### 問題 4: デバッグの困難性

バグを修正する場合：
```vba
' 「マイクロタスクの割当がおかしい」というバグ報告
' → 具体的な値で動作を確認したい
' → でも Excel を開かないと試せない
' → バグ報告者に「Excel で再現ファイルを送ってください」と依頼
' → ローカル環境がない人は修正不可
```

**理想的な修正方法**:
```typescript
// テストコードで再現
const result = runScheduling({
  tasks: [
    { duration: 0.1, assignee: 'Alice', ... },
    ...
  ],
  capacity: { Alice: 1.0, ... },
  holidays: [5, 12, 19],  // 5日、12日、19日が休日
});

// 修正前後で結果を比較
assert.deepEqual(result.grid[0][0], 0.1);  // 期待値
```

**VBA では不可能**: Excel がないと動かない

### リファクタリング後の改善

リファクタリングにより、ビジネスロジックが **I/O から分離** されます：

```vba
' TaskScheduler_Core.bas: 純粋な関数（Excel 不要）
Function CalcBaseStartIdx(currentLevel As Long, parentFinishIdx As Long, _
                          parentFinishAlloc As Double) As Long
    ' 入力: プリミティブ型のみ
    ' 出力: プリミティブ型のみ
    ' Excel オブジェクトへの参照なし
    ' → テスト可能！
End Function

Function ScheduleUnlockedTask(taskRow As Long, numDays As Long, ...) As Long
    ' 同様に、入力が配列＋プリミティブ型
    ' → テスト可能！
End Function

' TaskScheduler.bas: I/O だけを担当
Sub AutoScheduleTasks()
    ' フェーズ 0-3: I/O（Excel 読み込み）
    Call ReadConfigs(...)
    
    ' フェーズ 4-5: ビジネスロジック（Excel 不要）
    Call ScanLockedRows(...)
    Call ScheduleAllTasks(...)
    
    ' フェーズ 6: I/O（Excel 書き込み）
    Call WriteResults(...)
End Sub
```

**テスト可能になる**:
```typescript
// 単体テスト（Excel なし）
const result = CalcBaseStartIdx(2, 3, 0.3);
assert.equal(result, 3);

// 統合テスト（Mock Worksheet）
const mockApp = new MockApplication();
const ws = mockApp.Sheets('Data');
ws.setCellValue('A1', 'Task1');
ev.callProcedure('ScheduleUnlockedTask', [taskData, numDays, ...]);

// E2E テスト（実際の Excel）
// 必要に応じて実 Excel で検証
```

---

## 改善策：リファクタリング後のコード

### 改善 1: 関心の分離と構造化（UDT 導入） [[→ R-06](REFACTORING_TESTING_CATALOG.md#r-06) / [R-03](REFACTORING_TESTING_CATALOG.md#r-03)]

**リファクタリング前**:
```vba
Const COL_LEVEL_IDX As Long = 8
Const COL_OFFSET_IDX As Long = 9
Const COL_LOCK_IDX As Long = 13
Const COL_DURATION_IDX As Long = 15
Const COL_ASSIGNEE_IDX As Long = 17
Const CONFIG_ROW_START As Long = 8
Const CONFIG_ROW_END As Long = 14
Const CONFIG_COL_NAME As Long = 17
' ... バラバラに定義
```

**リファクタリング後**:
```vba
' TaskConfig: タスク関連の定数をグループ化
Type TaskConfig
    ROW_START As Long
    COL_LEVEL As Long
    COL_OFFSET As Long
    COL_LOCK As Long
    COL_DURATION As Long
    COL_ASSIGNEE As Long
    STR_LOCK_MARK As String
End Type

' CalendarConfig: カレンダー関連の定数をグループ化
Type CalendarConfig
    ROW_HEADER As Long
    ROW_HOLIDAY As Long
    COL_CALENDAR_START As Long
    STR_HOLIDAY_MARK As String
End Type

' AssigneeConfig: 担当者関連の定数をグループ化
Type AssigneeConfig
    COL_NAME As Long
    COL_LIMIT As Long
    ROW_START As Long
    ROW_END As Long
End Type
```

**メリット**:
- 定数が意味別に分類される → **概念が明確**
- パラメータ渡しが簡潔（`taskCfg` 1 つで全て）
- 変更が局所化（`TaskConfig` だけを変更すれば OK）

### 改善 2: Main Sub の簡潔化

**リファクタリング前** (SubProcedure 全体が 393 行):
```vba
Sub AutoScheduleTasks()
    ' Const（50 行）
    ' 変数宣言（30 行）
    ' Range 判定（20 行）
    ' データ読み込み（20 行）
    ' Config 読み込み（15 行）
    ' Phase 1: Locked Rows 処理（30 行）
    ' Phase 2: Scheduling（280 行）← ここが巨大！
    '   - Level ロジック
    '   - 依存関係計算
    '   - スケジューリングループ
    '   - Level Max 更新
    ' Write Back（5 行）
    ' Cleanup（10 行）
End Sub
```

**リファクタリング後** (メイン Sub が 203 行に短縮):
```vba
Sub AutoScheduleTasks()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    ' ===== 初期化フェーズ（読みやすい）=====
    Dim calCfg As CalendarConfig
    calCfg = InitCalendarConfig()         ' ← Config は別の関数で初期化
    Dim taskCfg As TaskConfig
    taskCfg = InitTaskConfig()
    Dim assigneeCfg As AssigneeConfig
    assigneeCfg = InitAssigneeConfig()
    
    ' Performance Settings
    Dim screenUpdateState As Boolean
    screenUpdateState = Application.ScreenUpdating
    Application.ScreenUpdating = False
    ' ...
    
    ' ===== Range 判定フェーズ（独立した関数） =====
    Dim lastRow As Long
    lastRow = GetLastTaskRow(ws, taskCfg)
    Dim lastCol As Long
    lastCol = GetLastCalendarCol(ws, calCfg)
    Dim numRows As Long
    numRows = CalcNumRows(lastRow, taskCfg)
    Dim numDays As Long
    numDays = CalcNumDays(lastCol, calCfg)
    
    ' ===== データ読み込みフェーズ（独立した関数） =====
    Dim rangeTask As Range
    Set rangeTask = GetTaskRange(ws, taskCfg, lastRow)
    Dim taskDataFrame As Variant
    taskDataFrame = rangeTask.Value
    
    Dim rangeSchedule As Range
    Set rangeSchedule = GetScheduleRange(ws, taskCfg, calCfg, lastRow, lastCol)
    Dim scheduleGrid As Variant
    scheduleGrid = rangeSchedule.Value
    
    ' ===== Phase 1: Locked Rows スキャン =====
    Call ScanLockedRows(taskCfg, numRows, numDays, taskDataFrame, scheduleGrid, assigneeUsage)
    
    ' ===== Phase 2: Scheduling ループ（明確な構造） =====
    Dim maxLevel As Long
    maxLevel = GetMaxLevel(taskDataFrame, numRows, taskCfg)
    Dim levelMaxFinish() As Long
    Dim levelMaxFinishAlloc() As Double
    ReDim levelMaxFinish(0 To maxLevel)
    ReDim levelMaxFinishAlloc(0 To maxLevel)
    
    For taskRow = 1 To numRows
        isLocked = IsRowLocked(taskDataFrame, taskRow, taskCfg)
        assigneeName = GetAssigneeName(taskDataFrame, taskRow, taskCfg)
        currentLevel = GetTaskLevel(taskDataFrame, taskRow, taskCfg)
        
        If currentLevel = 1 Then
            Erase levelMaxFinish
            Erase levelMaxFinishAlloc
        End If
        
        parentFinishIdx = levelMaxFinish(currentLevel - 1)
        parentFinishAlloc = levelMaxFinishAlloc(currentLevel - 1)
        baseStartIdx = CalcBaseStartIdx(currentLevel, parentFinishIdx, parentFinishAlloc)
        
        If isLocked Then
            Call FindLockedTaskFinish(...)
            Call UpdateLevelFinish(...)
        Else
            Call ScheduleUnlockedTask(...)
            Call UpdateLevelFinish(...)
        End If
    Next taskRow
    
    ' ===== Write Back =====
    rangeSchedule.Value = scheduleGrid
    
Cleanup:
    Application.ScreenUpdating = screenUpdateState
End Sub
```

**視認性の改善**:

```
| メイン Sub（203 行）
├─ InitXxxConfig() ........ 設定初期化（専門化）
├─ GetLastTaskRow() ....... Range 判定（専門化）
├─ GetTaskRange() ......... Range 構築（専門化）
├─ ScanLockedRows() ....... Phase 1（専門化）
├─ ScheduleUnlockedTask() . Phase 2（専門化）
└─ UpdateLevelFinish() .... 状態更新（専門化）
```

**読み方が自由**:
1. 全体構造を知りたい → Main Sub を読む（203 行で OK）
2. Phase 1 の詳細を知りたい → `ScanLockedRows()` を読む
3. スケジューリングロジックを知りたい → `ScheduleUnlockedTask()` を読む
4. Level 計算を知りたい → `CalcBaseStartIdx()` を読む

---

## 複雑度の分散

### リファクタリング前

```
複雑度分布
|
| ████████████████████████████████████ (393行が1つに集中)
|
+────────────────────────────────────
  Main Sub (393行)
```

**問題**: 
- 1 つの関数で全複雑度を処理
- 認知負荷が高い
- バグの可能性が高い

### リファクタリング後

```
複雑度分布
|
| ███ (Main Sub: 203行)
| ██ (ScanLockedRows: 20行)
| ██ (ScheduleUnlockedTask: 50行)
| █  (Helper Functions: 各 5-10行)
|
+────────────────────────────────────
```

**メリット**:
- 複雑度が分散
- 各関数の認知負荷が低い
- テストしやすい（1 つずつ検証可能）

---

## 実際の改善例：Locked Row 処理

### レガシーコード（30 行が Phase 1 に埋まっている）

```vba
' Phase 1: Scan Locked Rows （埋もれた）
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
```

**問題**:
- Phase 1 の処理が分からない（概要を読む必要）
- 「何をしているのか」が不明確
- デバッグ時に多くのコンテキストを保持する必要

### リファクタリング後（専門化された関数）

**メイン Sub**:
```vba
Call ScanLockedRows(taskCfg, numRows, numDays, taskDataFrame, scheduleGrid, assigneeUsage)
```

**TaskScheduler_Core.bas**:
```vba
Sub ScanLockedRows(taskCfg As TaskConfig, numRows As Long, numDays As Long, _
                   taskDataFrame As Variant, scheduleGrid As Variant, assigneeUsage As Object)
    Dim taskRow As Long, dayIdx As Long
    Dim assigneeName As String
    Dim newAllocArray() As Double
    
    For taskRow = 1 To numRows
        assigneeName = GetAssigneeName(taskDataFrame, taskRow, taskCfg)
        
        If assigneeName <> "" And IsRowLocked(taskDataFrame, taskRow, taskCfg) Then
            If Not assigneeUsage.Exists(assigneeName) Then
                ReDim newAllocArray(1 To numDays) As Double
                assigneeUsage.Add assigneeName, newAllocArray
            End If
            
            ' 負荷を集計
            Call AccumulateTaskAllocation(taskRow, numDays, scheduleGrid, assigneeUsage(assigneeName))
        End If
    Next taskRow
End Sub

' ヘルパー関数（責任が明確）
Sub AccumulateTaskAllocation(taskRow As Long, numDays As Long, scheduleGrid As Variant, allocArray() As Double)
    Dim dayIdx As Long
    Dim cellVal As Variant
    Dim existingAlloc As Double
    
    For dayIdx = 1 To numDays
        cellVal = scheduleGrid(taskRow, dayIdx)
        existingAlloc = 0
        If IsNumeric(cellVal) And Not IsEmpty(cellVal) Then
            existingAlloc = CDbl(cellVal)
        End If
        
        If existingAlloc > 0 Then
            allocArray(dayIdx) = allocArray(dayIdx) + existingAlloc
        End If
    Next dayIdx
End Sub
```

**メリット**:
1. **責任が明確** — ScanLockedRows は「Locked Row をスキャン」のみ
2. **テスト可能** — AccumulateTaskAllocation を単独でテスト
3. **変更が容易** — Locked Row の定義を変更？→ IsRowLocked() だけ変更
4. **再利用可能** — AccumulateTaskAllocation は他の場所からも呼び出し可能

---

## 複雑度が「分散」されることの利点

### ✅ メリット：複数の読み方が可能

**パターン 1: 全体構造を知りたい人**
```
1. Main Sub（203 行）を読む
   ↓ 3 分で全体像理解
```

**パターン 2: Phase 1 の詳細を知りたい人**
```
1. Main Sub の「Phase 1」セクション（1 行）
   ↓
2. ScanLockedRows() の実装（15 行）
   ↓ 1 分で完全理解
```

**パターン 3: スケジューリングロジックを知りたい人**
```
1. Main Sub の「Phase 2」セクション（20 行）
   ↓
2. ScheduleUnlockedTask() の実装（50 行）
   ↓
3. Micro-Task ロジック（20 行）
   ↓ 5 分で完全理解
```

**パターン 4: バグを修正したい人（変更箇所が明確）**
```
「Level 計算に bug がある」
  → CalcBaseStartIdx() だけを見る（5 行）
```

### ❌ レガシーコードの問題：読み方が 1 つ

```
全体を理解するには
  → すべての 393 行を順に読む必要がある
  ↓
「Locked Row の処理ってどこ？」
  → 250 行周辺を探す
  ↓
「でも Phase 1 と Phase 2 が混在している…」
  → 結局全体を追跡する羽目に
```

---

## 状態管理の分散

### レガシーコード（状態が共有される）

```
┌─────────────────────────────────┐
│ Sub AutoScheduleTasks()         │
├─────────────────────────────────┤
│ Dim metaData As Variant         │
│ Dim gridData As Variant         │◄─── Phase 1,2 で共有
│ Dim personUsage As Object       │
│ Dim levelMaxFinish() As Long    │
│ Dim capacityLimits As Object    │
├─────────────────────────────────┤
│ Phase 1: ロック行スキャン        │
│   metaData と gridData を読む     │
│   personUsage を更新             │
├─────────────────────────────────┤
│ Phase 2: スケジューリング        │
│   metaData, personUsage, 等を   │
│   同時に読み書き                 │
├─────────────────────────────────┤
│ Phase 3: Write Back             │
│   gridData をシートに書く        │
└─────────────────────────────────┘
```

**問題**:
- すべての状態がグローバル（Sub スコープ）
- Phase 1 と Phase 2 の依存関係が不明確
- デバッグ時に「どこで状態が狂ったのか」を追跡困難

### リファクタリング後（状態が局所化）

```
Main Sub
├─ Init Phase
│  ├─ calCfg ........ 初期化（使い回し）
│  ├─ taskCfg ....... 初期化（使い回し）
│  └─ assigneeCfg ... 初期化（使い回し）
│
├─ ScanLockedRows()
│  ├─ 入力: taskDataFrame, scheduleGrid
│  ├─ 出力: assigneeUsage（更新）
│  └─ 責任: Locked Row の負荷を集計
│
├─ Main Loop (Phase 2)
│  ├─ 入力: levelMaxFinish, levelMaxFinishAlloc（状態）
│  ├─ 各 Task で
│  │  ├─ GetTaskLevel()... Level を抽出
│  │  ├─ CalcBaseStartIdx()... 開始日を計算
│  │  └─ ScheduleUnlockedTask()... スケジュール
│  └─ 出力: scheduleGrid（更新）
│
└─ Write Back
   └─ scheduleGrid をシートに書く
```

**メリット**:
- 各関数の責任が明確
- 状態の流れが追跡可能
- テストがしやすい（状態を準備 → 関数呼び出し → 結果検証）

---

## チェックリスト：リファクタリングの効果測定

改善されたかどうかを判定：

- [ ] **読み方が複数ある** — Main Sub を読むだけで全体構造が分かるか？
- [ ] **複雑度が分散** — 各関数が 20 行程度か？
- [ ] **責任が明確** — 各関数の目的が関数名から分かるか？
- [ ] **パラメータが明確** — 「何が入力で、何が出力か」が関数シグネチャから分かるか？
- [ ] **テスト可能** — 1 つずつ単独でテスト（Excel シート依存なし）できるか？
- [ ] **変更が容易** — 要件変更時に修正箇所が明確か？

---

## 行数の比較

| 項目 | レガシー | リファクタリング後 |
|------|----------|------------------|
| Main Sub | 393 行 | 203 行（48%に短縮） |
| 定数定義 | 分散 | UDT で集約 |
| 関数数 | 1 | 20+ |
| 平均関数長 | 393 行 | 10-20 行 |
| 読むべき行数（全体理解） | 393 行 | 203 行（50%削減） |

---

## まとめ

### リファクタリングで何が改善されたのか

**表面的には「複雑になった」ように見えるが、実は：**

1. **複雑度が分散** — 1 つの巨大な関数が複数の小さな関数に分割
2. **読み方が自由** — 全体を読まなくても、興味のある部分だけ読める
3. **デバッグが容易** — 問題を特定しやすくなった
4. **テストが可能** — Excel シート依存なしで単体テスト可能
5. **変更が容易** — 要件変更時に修正範囲が明確

**結論**: 
```
「一見複雑に見えるが、実は理解しやすい」
= 複雑度が分散され、認知負荷が低い
```

このリファクタリングは、**REFACTORING_GUIDE.md で説明した「分割統治」と「関心の分離」の実践例**です。

---

## さらなるリファクタリング：セル位置の即値を避ける [[→ R-09](REFACTORING_TESTING_CATALOG.md#r-09)]

リファクタリング後のコードでも、列番号・行番号はソースに直書きされています。

```vba
Type TaskConfig
    COL_ASSIGNEE As Long  ' = 17（Q 列）
    ROW_START    As Long  ' = 19
    ' ...
End Type
```

この数値はシートに行や列を挿入・削除すると即座にずれ、バグの温床になります。

### 解決策：設定シートでセル位置を管理する

**シート「設定」** に Excel テーブル（名前: `設定`）を用意し、項目名と値の対で設定を管理します。

| 項目 | 値 |
|------|----|
| COL_ASSIGNEE | `=COLUMN(Q1)` |
| COL_DURATION | `=COLUMN(O1)` |
| ROW_TASK_START | `=ROW(A19)` |
| MAX_DAILY_LOAD | `1.0` |

セル位置には `=COLUMN(Q1)` や `=ROW(A19)` のようなセル参照式を入れておきます。シートに列・行を挿入・削除しても数式が自動で追従するため、設定テーブルの値が常に正しい位置を指し続けます。

### VBA 側の実装

`GetSetting` は [`sample/src/vba/LibSheet.bas`](../sample/src/vba/LibSheet.bas) に実装されています。シート名・テーブル名は埋め込まず、呼び出し元が `ListObject` を取得して渡します。

```vba
' LibSheet.bas
Function GetSetting(tbl As ListObject, itemName As String) As Variant
    Dim colItem As ListColumn
    Dim colValue As ListColumn
    Set colItem = tbl.ListColumns("項目")
    Set colValue = tbl.ListColumns("値")
    Dim i As Long
    For i = 1 To tbl.ListRows.Count
        If colItem.DataBodyRange(i, 1).Value = itemName Then
            GetSetting = colValue.DataBodyRange(i, 1).Value
            Exit Function
        End If
    Next i
    Err.Raise 1004, "GetSetting", "設定項目 '" & itemName & "' が見つかりません"
End Function

' InitTaskConfig() で設定テーブルを受け取り値を読み込む
Function InitTaskConfig(tbl As ListObject) As TaskConfig
    Dim cfg As TaskConfig
    cfg.COL_ASSIGNEE = GetSetting(tbl, "COL_ASSIGNEE")
    cfg.COL_DURATION = GetSetting(tbl, "COL_DURATION")
    cfg.ROW_START    = GetSetting(tbl, "ROW_TASK_START")
    InitTaskConfig = cfg
End Function

' 呼び出し元（TaskScheduler.bas）
Sub AutoScheduleTasks()
    Dim tbl As ListObject
    Set tbl = Worksheets("設定").ListObjects("設定")
    Dim taskCfg As TaskConfig
    taskCfg = InitTaskConfig(tbl)
    ' ...
End Sub
```

### メリット

| 比較 | ソース即値 | 設定シート管理 |
|------|-----------|--------------|
| 列挿入時の対応 | ソースを手修正 | 数式が自動追従 |
| 設定の一覧性 | Const が散在 | シートで一覧確認・変更可 |
| 非エンジニアによる調整 | 不可 | シートを直接編集できる |

### 代替案：名前付き範囲（Named Ranges）との比較

Excel の名前付き範囲（名前マネージャー）を使ってセル位置を管理する方法もありますが、以下のデメリットがあります。

| デメリット | 説明 |
|-----------|------|
| 定義が不可視 | 名前はワークブック内に隠れており、シートを見ただけでは何が定義されているか分からない |
| スコープが分かりづらい | ブック全体とシートローカルの2種類があり、意図せず混在すると参照が壊れる |
| シート操作で壊れやすい | シートの移動・コピー・改名で名前の参照が壊れることがあり、しかもエラーが出ないまま気づきにくい |

設定シート方式は設定の一覧がシート上に可視化されており、上記のリスクがないため、長期運用するブックでは設定シート方式を推奨します。

ビジネスロジックの変更なしにシートレイアウトを変えられるため、**長期運用するマクロほど効果が大きい**パターンです。
