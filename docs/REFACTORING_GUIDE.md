# VBA リファクタリングガイド：Excel オブジェクトとドメインロジックの分離

## はじめに

このドキュメントは、**巨大な Excel オブジェクト依存のレガシー VBA マクロを、テスト可能で保守性の高いコードに段階的にリファクタリングするための実践ガイド**です。

### なぜリファクタリングが必要なのか

**リファクタリング前の問題**:
```vba
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
```vba
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

```vba
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

```vba
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
```vba
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
```vba
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
```vba
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
```vba
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
```vba
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
```vba
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
```vba
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
```vba
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
```vba
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
```vba
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

---

## リファクタリングのステップバイステップガイド

### ステップ 1: ビジネスロジックを特定する

**質問**:
- この Sub の中で、**どの部分が Excel オブジェクトに依存しているか**？
- どの部分が **純粋な計算・ロジック** か？
- **どのデータが入力で、どのデータが出力** か？

**例**:
```vba
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

```vba
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

```vba
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

```vba
Sub TestCalculateDiscount()
    Debug.Assert CalculateDiscount(5000) = 500   ' 5000 * 0.10
    Debug.Assert CalculateDiscount(15000) = 3000 ' 15000 * 0.20
End Sub
```

### ステップ 5: Sub を I/O 専用に書き直す

```vba
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

```vba
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

```vba
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

```vba
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

```vba
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

```vba
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

```vba
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

```vba
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

```vba
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

```vba
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
- **`PLAN.md`** — Excel オブジェクト実装のロードマップ
