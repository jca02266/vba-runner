# VBA リファクタリングガイド：Excel オブジェクトとドメインロジックの分離

## はじめに

このドキュメントは、**巨大な Excel オブジェクト依存のレガシー VBA マクロを、テスト可能で保守性の高いコードに段階的にリファクタリングするための実践ガイド**です。

### なぜリファクタリングが必要なのか

**リファクタリング前の問題**:
```vb
' ❌ テスト不可能：Excel シートに依存
Sub CalculateMonthlySalesReport()
    Dim data As Variant
    Dim total As Long
    Dim i As Long
    
    ' シートから直接読む
    Set ws = Sheets("SalesData")
    data = ws.Range("A2:D1000").Value
    
    ' ビジネスロジック + I/O が混在
    total = 0
    For i = 1 To UBound(data)
        If data(i, 3) > 100 Then
            total = total + data(i, 4)
            ws.Range("E2").Offset(i - 1, 0).Value = "High"
        End If
    Next i
    
    ' 結果を直接シートに書く
    ws.Range("F2").Value = total
End Sub
```

**問題点**:
1. **テスト困難** — Excel シート依存のため、ユニットテスト不可能
2. **変更脆弱** — ロジック変更時に I/O 部分も影響を受ける
3. **デバッグ困難** — 複数の関心が混在しているため、問題の原因特定が難しい
4. **再利用不可能** — 他のシートやデータソースで同じロジックを適用できない

**リファクタリング後**:
```vb
' ✅ テスト可能：ドメインロジックは Excel に依存しない
Function CalculateTotalForHighValueItems(data() As Variant) As Long
    Dim total As Long, i As Long
    total = 0
    For i = 1 To UBound(data)
        If data(i, 3) > 100 Then
            total = total + data(i, 4)
        End If
    Next i
    CalculateTotalForHighValueItems = total
End Function

' I/O は別の Sub で担当
Sub CalculateMonthlySalesReport()
    ' 1. データを読む
    Dim data As Variant
    data = Sheets("SalesData").Range("A2:D1000").Value
    
    ' 2. ロジックを実行（テスト済み）
    Dim total As Long
    total = CalculateTotalForHighValueItems(data)
    
    ' 3. 結果を書く
    Sheets("SalesData").Range("F2").Value = total
End Sub
```

---

## 基本原則

### 原則 1: 職責の分離（Separation of Concerns） [[→ R-01](REFACTORING_TESTING_CATALOG.md#r-01)]

コードを 3 つの層に分ける：

```
┌─────────────────────────┐
│   I/O Layer (Sub)       │
│  - Sheets, Range        │
│  - Open, Close, Print   │
└────────────┬────────────┘
             │
┌────────────▼────────────┐
│   Business Logic        │
│   (Function)            │
│  - 計算・加工            │
│  - データ変換            │
│  - ビジネスルール        │
└────────────┬────────────┘
             │
┌────────────▼────────────┐
│   Data (Array/Variant)  │
│  - 素のデータ構造        │
│  - Excel 非依存         │
└─────────────────────────┘
```

**規則**:
- **Function** = ドメインロジック（入力 → 処理 → 出力）
- **Sub** = I/O 操作（Excel シート読み書き）
- **データ** = Variant 配列または構造化型（UDT）

### 原則 2: 入出力を関数の境界から排除

```vb
' ❌ 悪い例：関数内で I/O を行う
Function GetSalesTotal() As Long
    GetSalesTotal = Sheets("Sales").Range("B2").Value
End Function

' ✅ 良い例：データを引数で受け取る
Function GetSalesTotal(amount As Long) As Long
    ' ...計算...
End Function
```

### 原則 3: テスト可能な署名を設計する

```vb
' 関数は「入力データ」を受け取り「結果」を返すのみ
Function ProcessSalesData(salesData() As Variant, minAmount As Long) As Variant
    ' 実装...
End Function

' 複雑な結果は UDT で返す
Type SalesResult
    TotalAmount As Currency
    ItemCount As Long
    HighValueItems() As Variant
End Type

Function AnalyzeSales(salesData() As Variant) As SalesResult
    ' 実装...
End Function
```

---

## リファクタリングパターン集

### パターン 1: 単純な集計ロジック [[→ R-01](REFACTORING_TESTING_CATALOG.md#r-01)]

**リファクタリング前**（テスト不可能）:
```vb
Sub GenerateSalesReport()
    Dim ws As Worksheet
    Dim total As Currency
    Dim count As Long
    Dim i As Long
    
    Set ws = Sheets("Data")
    total = 0
    count = 0
    
    ' シートから直接読み取って集計
    For i = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If ws.Cells(i, 2).Value >= 1000 Then
            total = total + ws.Cells(i, 2).Value
            count = count + 1
        End If
    Next i
    
    ws.Range("F2").Value = total
    ws.Range("F3").Value = count
End Sub
```

**リファクタリング後**（テスト可能）:
```vb
' ✅ ドメインロジック（テスト対象）
Function SumHighValueSales(salesData() As Variant, threshold As Currency) As Currency
    Dim total As Currency
    Dim i As Long
    total = 0
    For i = 1 To UBound(salesData)
        If salesData(i, 2) >= threshold Then
            total = total + salesData(i, 2)
        End If
    Next i
    SumHighValueSales = total
End Function

Function CountHighValueSales(salesData() As Variant, threshold As Currency) As Long
    Dim count As Long
    Dim i As Long
    count = 0
    For i = 1 To UBound(salesData)
        If salesData(i, 2) >= threshold Then
            count = count + 1
        End If
    Next i
    CountHighValueSales = count
End Function

' I/O Sub（Excel 操作）
Sub GenerateSalesReport()
    ' 1. データ読み込み
    Dim ws As Worksheet
    Set ws = Sheets("Data")
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    Dim salesData As Variant
    salesData = ws.Range("A2:B" & lastRow).Value
    
    ' 2. ロジック実行
    Dim threshold As Currency
    threshold = 1000
    Dim total As Currency
    total = SumHighValueSales(salesData, threshold)
    Dim count As Long
    count = CountHighValueSales(salesData, threshold)
    
    ' 3. 結果書き込み
    ws.Range("F2").Value = total
    ws.Range("F3").Value = count
End Sub
```

**テスト**:
```vb
' ✅ テスト可能
Sub TestSumHighValueSales()
    Dim testData() As Variant
    testData = Array(Array(1, 500), Array(2, 1500), Array(3, 800), Array(4, 2000))
    
    Dim result As Currency
    result = SumHighValueSales(testData, 1000)
    
    Debug.Assert result = 3500  ' 1500 + 2000
End Sub
```

---

### パターン 2: データ加工と変換 [[→ R-01](REFACTORING_TESTING_CATALOG.md#r-01)]

**リファクタリング前**:
```vb
Sub FormatAndExportData()
    Dim ws As Worksheet
    Set ws = Sheets("RawData")
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    ' シートのデータを直接加工
    For i = 2 To lastRow
        Dim name As String
        name = ws.Cells(i, 1).Value
        Dim email As String
        email = ws.Cells(i, 2).Value
        
        ' データ加工ロジック + I/O が混在
        Dim formatted As String
        formatted = UCase(name) & " <" & LCase(email) & ">"
        
        ws.Cells(i, 3).Value = formatted
    Next i
End Sub
```

**リファクタリング後**:
```vb
' ✅ ドメインロジック
Function FormatUserRecord(name As String, email As String) As String
    FormatUserRecord = UCase(name) & " <" & LCase(email) & ">"
End Function

Function FormatUserData(rawData() As Variant) As Variant()
    Dim result() As Variant
    ReDim result(1 To UBound(rawData))
    Dim i As Long
    For i = 1 To UBound(rawData)
        result(i) = FormatUserRecord(rawData(i, 1), rawData(i, 2))
    Next i
    FormatUserData = result
End Function

' I/O Sub
Sub FormatAndExportData()
    ' 1. データ読み込み
    Dim ws As Worksheet
    Set ws = Sheets("RawData")
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    Dim rawData() As Variant
    rawData = ws.Range("A2:B" & lastRow).Value
    
    ' 2. ロジック実行
    Dim formatted() As Variant
    formatted = FormatUserData(rawData)
    
    ' 3. 結果書き込み
    Dim i As Long
    For i = 1 To UBound(formatted)
        ws.Cells(i + 1, 3).Value = formatted(i)
    Next i
End Sub
```

---

### パターン 3: 複雑なビジネスロジック（UDT を使用） [[→ R-01](REFACTORING_TESTING_CATALOG.md#r-01) / [R-03](REFACTORING_TESTING_CATALOG.md#r-03)]

**リファクタリング前**:
```vb
Sub CalculateCommission()
    Dim ws As Worksheet
    Set ws = Sheets("Sales")
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    ' 複雑なロジックが Sub に詰まっている
    For i = 2 To lastRow
        Dim salesAmount As Currency
        salesAmount = ws.Cells(i, 2).Value
        
        Dim region As String
        region = ws.Cells(i, 3).Value
        
        Dim commission As Currency
        ' ビジネスルール + I/O が混在
        If region = "Tokyo" Then
            If salesAmount >= 1000000 Then
                commission = salesAmount * 0.15
            Else
                commission = salesAmount * 0.10
            End If
        ElseIf region = "Osaka" Then
            If salesAmount >= 500000 Then
                commission = salesAmount * 0.12
            Else
                commission = salesAmount * 0.08
            End If
        End If
        
        ws.Cells(i, 4).Value = commission
    Next i
End Sub
```

**リファクタリング後**（UDT でドメインモデルを定義）:
```vb
' ドメインモデル
Type SalesRecord
    amount As Currency
    region As String
End Type

' ✅ ビジネスロジック：純粋な関数
Function CalculateCommissionRate(amount As Currency, region As String) As Double
    Select Case region
        Case "Tokyo"
            If amount >= 1000000 Then
                CalculateCommissionRate = 0.15
            Else
                CalculateCommissionRate = 0.10
            End If
        Case "Osaka"
            If amount >= 500000 Then
                CalculateCommissionRate = 0.12
            Else
                CalculateCommissionRate = 0.08
            End If
        Case Else
            CalculateCommissionRate = 0.05
    End Select
End Function

Function CalculateCommission(amount As Currency, region As String) As Currency
    CalculateCommission = amount * CalculateCommissionRate(amount, region)
End Function

' I/O Sub
Sub CalculateCommission()
    ' 1. データ読み込み
    Dim ws As Worksheet
    Set ws = Sheets("Sales")
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    Dim salesData() As Variant
    salesData = ws.Range("A2:C" & lastRow).Value
    
    ' 2. 結果計算
    Dim i As Long
    For i = 1 To UBound(salesData)
        Dim commission As Currency
        commission = CalculateCommission(salesData(i, 2), salesData(i, 3))
        ws.Cells(i + 1, 4).Value = commission
    Next i
End Sub
```

**テスト**:
```vb
Sub TestCalculateCommission()
    ' Tokyo region, high amount
    Debug.Assert CalculateCommission(1500000, "Tokyo") = 225000  ' 1500000 * 0.15
    
    ' Tokyo region, low amount
    Debug.Assert CalculateCommission(500000, "Tokyo") = 50000     ' 500000 * 0.10
    
    ' Osaka region, high amount
    Debug.Assert CalculateCommission(600000, "Osaka") = 72000     ' 600000 * 0.12
End Sub
```

---

### パターン 4: 状態を持つ処理（オブジェクト指向版） [[→ R-04](REFACTORING_TESTING_CATALOG.md#r-04)]

**リファクタリング前**（グローバル変数を使用）:
```vb
' ❌ グローバル状態（テスト困難）
Dim totalProcessed As Long
Dim totalErrors As Long

Sub ProcessBatchData()
    totalProcessed = 0
    totalErrors = 0
    
    Dim ws As Worksheet
    Set ws = Sheets("Batch")
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    For i = 2 To lastRow
        On Error GoTo ErrorHandler
        
        Dim value As Long
        value = ws.Cells(i, 1).Value
        
        If value > 0 Then
            ws.Cells(i, 2).Value = value * 2
            totalProcessed = totalProcessed + 1
        Else
            totalErrors = totalErrors + 1
        End If
        
        On Error GoTo 0
    Next i
    
    ws.Range("E2").Value = totalProcessed
    ws.Range("E3").Value = totalErrors
    Exit Sub
ErrorHandler:
    totalErrors = totalErrors + 1
End Sub
```

**リファクタリング後**（Class を使用）:
```vb
' ✅ ドメインモデル：Class で状態を管理
Class BatchProcessor
    Private pProcessed As Long
    Private pErrors As Long
    
    Sub Process(value As Long)
        If value > 0 Then
            pProcessed = pProcessed + 1
        Else
            pErrors = pErrors + 1
        End If
    End Sub
    
    Function GetProcessed() As Long
        GetProcessed = pProcessed
    End Function
    
    Function GetErrors() As Long
        GetErrors = pErrors
    End Function
    
    Function GetTotal() As Long
        GetTotal = pProcessed + pErrors
    End Function
End Class

' テスト可能なロジック
Sub TestBatchProcessor()
    Dim processor As New BatchProcessor
    processor.Process(100)
    processor.Process(50)
    processor.Process(-1)
    processor.Process(0)
    
    Debug.Assert processor.GetProcessed() = 2
    Debug.Assert processor.GetErrors() = 2
    Debug.Assert processor.GetTotal() = 4
End Sub

' I/O Sub
Sub ProcessBatchData()
    ' 1. オブジェクト生成
    Dim processor As New BatchProcessor
    
    ' 2. データ読み込みと処理
    Dim ws As Worksheet
    Set ws = Sheets("Batch")
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    Dim i As Long
    For i = 2 To lastRow
        Dim value As Long
        value = ws.Cells(i, 1).Value
        
        processor.Process(value)
        
        If value > 0 Then
            ws.Cells(i, 2).Value = value * 2
        End If
    Next i
    
    ' 3. 結果書き込み
    ws.Range("E2").Value = processor.GetProcessed()
    ws.Range("E3").Value = processor.GetErrors()
End Sub
```

### パターン 5: テーブル駆動パターン [[→ R-11](REFACTORING_TESTING_CATALOG.md#r-11)]

**概要**: 大量の同じ構造を持つ分岐を、**データテーブルとシンプルなルックアップロジック**に置き換える。複雑な if-else-if チェーンが繰り返されている場合に特に有効。

**リファクタリング前**（ネストされた分岐地獄）:
```vb
' ❌ 71行の分岐地獄：5部門 × 4閾値 = 20パターン
Function GetApprover(amount As Long, department As String) As String
    If department = "Sales" Then
        If amount < 50000 Then
            GetApprover = "Manager"
        ElseIf amount < 500000 Then
            GetApprover = "Director"
        ElseIf amount < 2000000 Then
            GetApprover = "VP"
        Else
            GetApprover = "CFO"
        End If
    ElseIf department = "Marketing" Then
        If amount < 30000 Then
            GetApprover = "Manager"
        ElseIf amount < 300000 Then
            GetApprover = "Director"
        ' ... 同じパターンが5部門分繰り返される
    End If
End Function
```

**リファクタリング後**（テーブル駆動）:
```vb
' ✅ 型定義：テーブル行
Type ApprovalRule
    department As String
    threshold1 As Long
    approver1 As String
    threshold2 As Long
    approver2 As String
    threshold3 As Long
    approver3 As String
    defaultApprover As String
End Type

Dim g_rules() As ApprovalRule

Sub InitializeApprovalRules()
    ReDim g_rules(4)
    
    g_rules(0).department = "Sales"
    g_rules(0).threshold1 = 50000
    g_rules(0).approver1 = "Manager"
    g_rules(0).threshold2 = 500000
    g_rules(0).approver2 = "Director"
    ' ... テーブルをデータで埋める
End Sub

' ✅ シンプルなルックアップロジック（23行に削減）
Function GetApprover(amount As Long, department As String) As String
    Dim i As Integer
    Dim rule As ApprovalRule
    
    For i = LBound(g_rules) To UBound(g_rules)
        rule = g_rules(i)
        If rule.department = department Then
            If amount < rule.threshold1 Then
                GetApprover = rule.approver1
            ElseIf amount < rule.threshold2 Then
                GetApprover = rule.approver2
            ElseIf amount < rule.threshold3 Then
                GetApprover = rule.approver3
            Else
                GetApprover = rule.defaultApprover
            End If
            Exit Function
        End If
    Next i
    
    GetApprover = "Unknown"
End Function
```

**テスト**:
```vb
Sub TestApprovalRules()
    ' テーブルの初期化
    InitializeApprovalRules
    
    ' テストケース：Sales部, 100,000円 → Director
    Debug.Assert GetApprover(100000, "Sales") = "Director"
    
    ' Marketing部, 15,000円 → Manager
    Debug.Assert GetApprover(15000, "Marketing") = "Manager"
    
    ' 不明な部門 → Unknown
    Debug.Assert GetApprover(100000, "Unknown") = "Unknown"
End Sub
```

**パターン 5 の利点**:
- **コード削減**: 71行 → 47行（35%削減、規則が多いほど効果大）
- **保守性向上**: ルール変更 = **コード編集ではなくデータ更新**
- **変更頻度が高い場合に有効**: 金額閾値や承認者が頻繁に変わる環境
- **外部データ連携**: テーブルを Excel シートや JSON から読み込み可能

**パターン 5 の注意点**:
- テーブル行の構造が固定化する（新しい条件型が追加できない）
- Dictionary ベースの動的アプローチ（パターン 5b）で解決可能

**パターン 5b: 動的テーブル（Dictionary）**:
```vb
' 外部データから動的にルールを構築
Function BuildApprovalRulesDictionary() As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    Dim salesRules As Object
    Set salesRules = CreateObject("Scripting.Dictionary")
    salesRules.Add "levels", Array( _
        CreateLevelTuple(50000, "Manager"), _
        CreateLevelTuple(500000, "Director"))
    dict.Add "Sales", salesRules
    
    Set BuildApprovalRulesDictionary = dict
End Function

Function GetApproverFromTable(amount As Long, department As String, rules As Object) As String
    If Not rules.Exists(department) Then
        GetApproverFromTable = "Unknown"
        Exit Function
    End If
    
    Dim deptRules As Object
    Set deptRules = rules(department)
    Dim levels As Variant
    levels = deptRules("levels")
    
    Dim i As Integer
    For i = LBound(levels) To UBound(levels)
        If amount < levels(i)("threshold") Then
            GetApproverFromTable = levels(i)("approver")
            Exit Function
        End If
    Next i
    
    GetApproverFromTable = deptRules("default")
End Function
```

**テーブル駆動の検出と提案**:

IDE の自動検出機能（`TableDrivenDetector`）により、リファクタリング候補を自動判定可能：

| 判定項目 | 条件 |
|---------|------|
| **外側分岐数** | ≥ 3 個（異なる条件値） |
| **内側分岐数** | ≥ 3 個（反復するパターン） |
| **形状一致** | すべての分岐が同じ構造 |
| **代入パターン** | リテラル/識別子のみ（複雑な式なし） |
| **副作用** | 関数呼び出しがない |
| **信頼度** | ≥ 70/100 → 強く推奨 |

**テーブル設計の自動提案**（将来の拡張）:
```
検出結果：
  - 5個の外側分岐（department）
  - 3個の内側分岐（amount thresholds）
  
自動提案テーブル設計：
  Type ApprovalRule
    department As String       ← 外側分岐のキー
    threshold1 As Long         ← 内側分岐①の条件値
    approver1 As String        ← 内側分岐①の結果
    threshold2 As Long         ← 内側分岐②の条件値
    approver2 As String        ← 内側分岐②の結果
    defaultApprover As String  ← else節の結果
  End Type
```

**テストケース生成の可能性**（将来の拡張）:

リファクタリング前後の動作検証テストを自動生成：
```vb
' 自動生成テスト：カバレッジ = 5部門 × (3閾値 + 1デフォルト) = 20パターン

Sub TestApprovalRulesAutoGenerated()
    InitializeApprovalRules
    
    ' Sales部テスト
    Debug.Assert GetApprover(10000, "Sales") = "Manager"      ' < threshold1
    Debug.Assert GetApprover(100000, "Sales") = "Director"    ' < threshold2
    Debug.Assert GetApprover(1000000, "Sales") = "VP"         ' < threshold3
    Debug.Assert GetApprover(5000000, "Sales") = "CFO"        ' >= threshold3
    
    ' Marketing部テスト（異なる閾値で同じパターン）
    Debug.Assert GetApprover(10000, "Marketing") = "Manager"
    Debug.Assert GetApprover(100000, "Marketing") = "Director"
    
    ' ... 全部門カバレッジ自動生成
End Sub
```

---

### パターン 6: 副作用が絡み合った関数のリファクタリング [[→ R-12](REFACTORING_TESTING_CATALOG.md#r-12) / [R-13](REFACTORING_TESTING_CATALOG.md#r-13)]

**概要**: `Sheets` / `MsgBox` / ファイル I/O などの副作用がロジックの途中に散在し、テストも書けない大きな関数に対処する2つのパターン。副作用の性質に応じてどちらを使うか選ぶ。

---

#### パターン 6a: FC/IS（副作用を末尾にまとめられる場合）

副作用の実行タイミングを後ろにまとめても結果が変わらない場合に、副作用なしの純粋ロジック（Core）と副作用だけを実行する外殻（Shell）に分離する。

**Before: 副作用がロジックの途中に散在**

```vb
Sub ProcessApproval(dept As String, amount As Long)
    Dim decision As String
    If dept = "Engineering" And amount < 100000 Then
        decision = "TeamLead"
    ElseIf dept = "Sales" Then
        decision = "Manager"
    End If

    WriteApprovalLog dept, decision        ' 副作用①（ロジックの途中）

    If decision = "TeamLead" Then amount = amount * 0.9

    Sheets("Result").Cells(1, 1).Value = decision  ' 副作用②
End Sub
```

**After: 副作用の実行順を変えても安全であることを確認してから分離**

```vb
' Core: 純粋関数。副作用なし
Function DecideApproval(dept As String, amount As Long) As String
    If dept = "Engineering" And amount < 100000 Then
        DecideApproval = "TeamLead"
    ElseIf dept = "Sales" Then
        DecideApproval = "Manager"
    End If
End Function

' Shell: Core を呼び、副作用をまとめて実行するだけ
Sub ProcessApproval(dept As String, amount As Long)
    Dim decision As String
    decision = DecideApproval(dept, amount)         ' Core
    WriteApprovalLog dept, decision                 ' 副作用①
    Sheets("Result").Cells(1, 1).Value = decision   ' 副作用②
End Sub
```

Core（`DecideApproval`）は Excel なしでユニットテストできる。Shell は手動確認または統合テストで検証する。

**注意**: 副作用①の実行タイミングを後ろに移動しても本当に問題ないか慎重に確認すること（ログの目的・失敗時の挙動など）。Core に `ByRef` を使った変形が必要になる場合は FC/IS の適用自体を再検討し、パターン 6b（Seam）を選ぶ。

---

#### パターン 6b: Seam（副作用の結果がロジックに影響する場合）

副作用の出力が後続のロジック分岐に影響する、または副作用への入力値が処理途中の計算結果に依存するため、後ろにまとめられない場合に使う。**テストが書ける状態を作るための前段階**として副作用の塊を関数抽出（[R-01](REFACTORING_TESTING_CATALOG.md#r-01)）で関数（縫い目）として抽出し、テスト時にモックに差し替えられるようにする。

**Before: 副作用の出力がロジックに入り込んでいる**

```vb
' 副作用①の結果で分岐 → 中間計算 → 副作用②の入力に使う → 副作用②の結果で分岐
' 副作用②への入力 route は副作用①の出力から計算されるため、引数に追い出しても呼び出し元に問題が移るだけ
Function ProcessApproval(dept As String, amount As Long) As String
    Dim status As String
    status = Sheets("Config").Cells(GetDeptRow(dept), 2).Value  ' 副作用①
    If status <> "active" Then
        ProcessApproval = "Rejected": Exit Function
    End If

    Dim route As String
    route = DecideRoute(dept, status)   ' 中間計算（①の出力に依存）

    Dim limit As Long
    limit = Sheets("Limits").Cells(GetRouteRow(route), 2).Value  ' 副作用②（route を入力に使う）
    If amount > limit Then
        ProcessApproval = "Escalate"
    Else
        ProcessApproval = "Approve"
    End If
    Sheets("Log").Cells(GetNextLogRow(), 1).Value = ProcessApproval  ' 副作用③
End Function
```

**Step 1: 副作用の塊を縫い目（関数）として関数抽出（[R-01](REFACTORING_TESTING_CATALOG.md#r-01)）する**

```vb
Function ProcessApproval(dept As String, amount As Long) As String
    Dim status As String
    status = ReadDeptStatus(dept)              ' ← 縫い目①
    If status <> "active" Then
        ProcessApproval = "Rejected": Exit Function
    End If
    Dim route As String
    route = DecideRoute(dept, status)          ' 中間計算（純粋ロジック）
    Dim limit As Long
    limit = ReadApprovalLimit(route)           ' ← 縫い目②（route を受け取る）
    If amount > limit Then
        ProcessApproval = "Escalate"
    Else
        ProcessApproval = "Approve"
    End If
    WriteApprovalLog dept, ProcessApproval     ' ← 縫い目③
End Function

Function ReadDeptStatus(dept As String) As String
    ReadDeptStatus = Sheets("Config").Cells(GetDeptRow(dept), 2).Value
End Function

Function ReadApprovalLimit(route As String) As Long
    ReadApprovalLimit = Sheets("Limits").Cells(GetRouteRow(route), 2).Value
End Function

Sub WriteApprovalLog(dept As String, result As String)
    Sheets("Log").Cells(GetNextLogRow(), 1).Value = result
End Sub
```

**Step 2: Excel で実際に動かしてスナップショットを記録する**（VBA Runner は `Sheets` をスタブするため不可）

**Step 3: TypeScript テストで縫い目をモックに差し替えてロジックをテスト**

```typescript
const runner = new VBARunner("ProcessApproval.bas");
runner.mock("ReadDeptStatus",    (_dept: string)  => "active");
runner.mock("ReadApprovalLimit", (_route: string) => 500000);
const logSpy = runner.spy("WriteApprovalLog");

assert.strictEqual(runner.run("ProcessApproval", ["Engineering", 100000]), "Approve");
assert.strictEqual(logSpy.calls[0].args, ["Engineering", "Approve"]);

// 上限超えのケース
runner.mock("ReadApprovalLimit", (_route: string) => 50000);
assert.strictEqual(runner.run("ProcessApproval", ["Engineering", 100000]), "Escalate");
```

縫い目が作れると `ReadDeptStatus` / `ReadApprovalLimit` の戻り値を自由に設定でき、ロジックの全パターンをテストできる。

**モックの粒度**: 低レベル API（`Sheets`・`MsgBox`）ではなくビジネス的な意味を持つ塊単位で縫い目を作る。塊の中に `ByRef` 変数への更新が含まれる場合はモックでもその更新を再現する。

**重要**: 縫い目を作る関数抽出（[R-01](REFACTORING_TESTING_CATALOG.md#r-01)）自体はテストなしで行うため、抽出後は必ず Excel 上での手動実行で動作確認してからスナップショット（[T-13](REFACTORING_TESTING_CATALOG.md#t-13)）を記録する。

---

#### パターン選択の判断基準

| 状況 | 選択 |
|---|---|
| 副作用の実行順を後ろに変えても結果が変わらない | パターン 6a（FC/IS） |
| 副作用の結果が後続のロジック分岐に影響する | パターン 6b（Seam） |
| 副作用への入力が処理途中の計算結果に依存する | パターン 6b（Seam） |
| Core に `ByRef` が必要になる変形が生じる | FC/IS を断念して 6b（Seam） |

---

## リファクタリングのステップバイステップガイド

### ステップ 1: ビジネスロジックを特定する

**質問**:
- この Sub の中で、**どの部分が Excel オブジェクトに依存しているか**？
- どの部分が **純粋な計算・ロジック** か？
- **どのデータが入力で、どのデータが出力** か？

**例**:
```vb
Sub CalculateDiscount()
    ' [I/O] シートから読む
    Dim price As Currency
    price = Sheets("Config").Range("A1").Value
    
    ' [LOGIC] 割引を計算
    Dim discountRate As Double
    If price >= 10000 Then
        discountRate = 0.20
    Else
        discountRate = 0.10
    End If
    Dim discount As Currency
    discount = price * discountRate
    
    ' [I/O] シートに書く
    Sheets("Result").Range("B1").Value = discount
End Sub
```

### ステップ 2: ロジック部分を Function に抽出

```vb
' ✅ 抽出されたロジック（Excel 非依存）
Function CalculateDiscount(price As Currency) As Currency
    Dim discountRate As Double
    If price >= 10000 Then
        discountRate = 0.20
    Else
        discountRate = 0.10
    End If
    CalculateDiscount = price * discountRate
End Function
```

### ステップ 3: パラメーターを最小化する

ロジック Function が依存する **すべてのデータを引数** にする：

```vb
' ❌ Excel に依存
Function CalculateDiscount() As Currency
    Dim price As Currency
    price = Sheets("Config").Range("A1").Value  ' ← Excel 依存
    ' ...
End Function

' ✅ 完全に独立
Function CalculateDiscount(price As Currency) As Currency
    ' ...
End Function
```

### ステップ 4: テストを書く

```vb
Sub TestCalculateDiscount()
    Debug.Assert CalculateDiscount(5000) = 500   ' 5000 * 0.10
    Debug.Assert CalculateDiscount(15000) = 3000 ' 15000 * 0.20
End Sub
```

### ステップ 5: Sub を I/O 専用に書き直す

```vb
Sub CalculateDiscount()
    ' 1. 読み込み
    Dim price As Currency
    price = Sheets("Config").Range("A1").Value
    
    ' 2. 計算（テスト済みロジック）
    Dim discount As Currency
    discount = CalculateDiscount(price)
    
    ' 3. 書き込み
    Sheets("Result").Range("B1").Value = discount
End Sub
```

---

## よくある間違いと対策

### 間違い 1: ロジック Function に Excel オブジェクトを含める

```vb
' ❌ ダメ：Function が Sheets に依存している
Function CalculateTotal() As Currency
    Dim ws As Worksheet
    Set ws = Sheets("Data")  ' ← Excel 依存
    ' ...
End Function

' ✅ 正解：データを引数で受け取る
Function CalculateTotal(data() As Variant) As Currency
    ' ...
End Function
```

### 間違い 2: 複数のデータソースから読む

```vb
' ❌ ダメ：複数のシートに依存
Sub ProcessData()
    Dim data1 As Variant
    data1 = Sheets("Sheet1").Range("A:A").Value
    Dim data2 As Variant
    data2 = Sheets("Sheet2").Range("B:B").Value
    ' ...
End Sub

' ✅ 正解：I/O 層で統合してからロジック層に渡す
Sub ProcessData()
    Dim data1 As Variant
    data1 = Sheets("Sheet1").Range("A:A").Value
    Dim data2 As Variant
    data2 = Sheets("Sheet2").Range("B:B").Value
    
    ' 統合後、ロジック層に渡す
    Dim result As Variant
    result = AnalyzeCombinedData(data1, data2)
    
    ' 結果を書く
    Sheets("Result").Range("C:C").Value = result
End Sub

' ✅ ロジック層：Excel 非依存
Function AnalyzeCombinedData(data1() As Variant, data2() As Variant) As Variant
    ' ...
End Function
```

### 間違い 3: 戻り値の型が曖昧

```vb
' ❌ ダメ：Variant（何が返るか不明確）
Function Process() As Variant
    ' ...
End Function

' ✅ 正解：具体的な型を指定
Function Process() As Long
    ' ...
End Function

' 複数の結果が必要な場合は UDT を使用
Type ProcessResult
    Total As Long
    Count As Long
    Average As Double
End Type

Function Process(data() As Variant) As ProcessResult
    ' ...
End Function
```

### 間違い 4: ロジックと I/O のサイズが 1:1

```vb
' ❌ ダメ：大量の I/O コードの中に少しのロジック
Sub GenerateReport()
    ' ← 100 行の I/O 操作
    ws.Range(...).Value = ...
    ws.Range(...).Value = ...
    ' ← 1 行のロジック
    If x > 10 Then y = z
    ' ← 100 行の I/O 操作
    ws.Range(...).Value = ...
End Sub

' ✅ 正解：ロジックを集約する
Function CalculateReport(data() As Variant) As Variant()
    ' 50 行のロジック
    ' ...
    CalculateReport = result
End Function

Sub GenerateReport()
    ' 10 行：読む
    Dim data As Variant
    data = ws.Range(...).Value
    
    ' 1 行：計算
    Dim result As Variant
    result = CalculateReport(data)
    
    ' 10 行：書く
    ws.Range(...).Value = result
End Sub
```

---

## 段階的なリファクタリング戦略

### フェーズ 1: ホットスポットを特定 [[→ R-07](REFACTORING_TESTING_CATALOG.md#r-07)]

**最初にすべてをリファクタリングしない**。以下を優先：

1. **テストが必要な部分** — ビジネスロジック
2. **頻繁に変更される部分** — 要件変更が多い
3. **複雑な部分** — バグが多い

```vb
' 優先度：高（複雑なビジネスルール）
Function CalculateCommission(sales As Currency, region As String) As Currency
    ' 複雑で頻繁に変更される
End Function

' 優先度：低（単純な I/O）
Sub ExportToCSV()
    ' 単純で変更が少ない
End Sub
```

### フェーズ 2: 小さく始める [[→ R-08](REFACTORING_TESTING_CATALOG.md#r-08)]

複数の Function を同時にリファクタリングしない。1つずつ：

```vb
' Step 1: 1 つの Function を抽出
Function GetTotal(data() As Variant) As Long
    ' ...
End Function

' Step 2: その Function をテスト
Sub TestGetTotal()
    Debug.Assert GetTotal(Array(1, 2, 3)) = 6
End Sub

' Step 3: Sub から呼び出し
Sub Main()
    Dim result As Long
    result = GetTotal(...)
    ws.Range("A1").Value = result
End Sub
```

### フェーズ 3: 段階的な導入 [[→ R-08](REFACTORING_TESTING_CATALOG.md#r-08)]

既存コードと新しいコードを共存させる：

```vb
' ✅ 新しいロジック（テスト済み）
Function NewCalculateTotal(data() As Variant) As Long
    ' ...
End Function

' 旧コード（レガシー）
Sub OldGenerateReport()
    ' 古い実装をそのまま保持
End Sub

' 移行用コード
Sub Main()
    ' 新しいロジックを使う
    Dim result As Long
    result = NewCalculateTotal(data)
    ws.Range("A1").Value = result
    
    ' 旧コードはまだ残す（レグレッション対策）
End Sub
```

---

## リファクタリングのチェックリスト

### ロジック Function の設計

- [ ] **入力** — すべてのパラメーターが Function の引数か？
- [ ] **出力** — 戻り値の型が明確か？（Variant ではない）
- [ ] **依存性** — Excel オブジェクト（Sheets, Range）に依存していないか？
- [ ] **副作用** — Debug.Print や MsgBox を含んでいないか？（テスト時にエラーになる）

### I/O Sub の設計

- [ ] **責任** — 読む、計算、書くの 3 ステップが明確か？
- [ ] **テスト対象外** — ロジックをテストする必要がないか？
- [ ] **エラーハンドリング** — I/O 失敗時の処理があるか？

### テストの実装

- [ ] **基本系** — 通常ケースをテストしているか？
- [ ] **エッジケース** — 境界値をテストしているか？
- [ ] **エラー系** — 異常系をテストしているか？

### ドキュメント

- [ ] **ビジネスルール** — ロジック Function に何を計算しているか書いているか？
- [ ] **前提条件** — データの形式や制約を明記しているか？
- [ ] **戻り値** — 戻り値が何を意味しているか説明しているか？

---

## よくある質問（FAQ）

### Q1: Excel 操作なしで VBA コードをテストできるのか？

**A**: はい。ドメインロジックを Function に分離すれば、配列データでテストできます。

```vb
' シートを使わずにテスト
Sub TestCalculateTax()
    Dim testData() As Variant
    testData = Array(100, 200, 300)
    
    Dim tax As Long
    tax = CalculateTax(testData)
    
    Debug.Assert tax = 60  ' (100 + 200 + 300) * 0.1
End Sub
```

### Q2: 既存のレガシーコード全体をリファクタリングする必要があるのか？

**A**: いいえ。テストが必要な部分から段階的に進めてください。

### Q3: Class（オブジェクト指向）を使うべきか？

**A**: 以下の場合に検討：
- 状態を持つ処理（複数回呼び出しで値が変わる）
- 関連する複数の Function がある
- ドメインモデルとして表現したい

単純な計算なら **Function だけで十分** です。

### Q4: リファクタリング後、元のコードと互換性を保つ必要があるのか？

**A**: 公開 API（他のマクロから呼ばれる Sub）は、署名を変えないようにしてください。内部実装は自由に変更できます。

```vb
' 公開 API：変更しない
Sub Main()
    ' ...既存のコードから呼ばれている...
End Sub

' 内部実装：大幅リファクタリングOK
Sub Main()
    ' ✅ 新しい実装に変更
    Dim data As Variant
    data = ReadData()
    Dim result As Variant
    result = AnalyzeData(data)
    WriteResult(result)
End Sub
```

---

## 関連ドキュメント

- **`TESTING_STRATEGY.md`** — VBA テストの設計原則
- **`TEST_FRAMEWORK_GUIDE.md`** — JavaScript テストフレームワークの活用
