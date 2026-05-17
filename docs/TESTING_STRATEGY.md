# VBA マクロ単体テスト戦略

## はじめに

このドキュメントは、**VBA Runner 上で実行可能なマクロの単体テストを効果的に行うための設計原則とパターン**を説明します。

### VBA Runner の制約

- **Excel がない** (Node.js / ブラウザ環境)
- **本物の Excel オブジェクトモデルは使用できない**
- Excel オブジェクト(`Sheets`, `Range`等) は VBA Runner で自前実装する必要がある

この制約の中で、**テスト可能で保守性の高い VBA コードを書く方法**を提案します。

---

## 原則1: Domain Logic と Excel I/O の徹底的分離

### 問題のあるコード（テストしづらい）

```vba
' ❌ Excel オブジェクトに直接依存
Function CalculateSalesTotal() As Long
    Dim data As Variant
    data = Sheets("SalesData").Range("A1:A5").Value
    Dim total As Long
    Dim i As Integer
    For i = LBound(data) To UBound(data)
        total = total + data(i)
    Next i
    CalculateSalesTotal = total
End Function

Sub Main()
    Dim result As Long
    result = CalculateSalesTotal()
    Sheets("Report").Range("C1").Value = result
End Sub
```

**問題**:
- `CalculateSalesTotal()` をテストするには Excel オブジェクトの mock が必要
- mock が複雑 → テストコードが本番コードより複雑
- 純粋な計算ロジックが Excel I/O に埋没

### 改善されたコード（テスト可能）

```vba
' === Domain Logic 層 ===
' 入力：配列（プリミティブなデータ）
' 出力：計算結果
' 依存：なし（Excel I/O なし）
Function CalculateTotalFromArray(dataArray As Variant) As Long
    Dim total As Long
    Dim i As Integer
    For i = LBound(dataArray) To UBound(dataArray)
        total = total + dataArray(i)
    Next i
    CalculateTotalFromArray = total
End Function

' === Excel I/O 層 ===
' 責務：Excel からデータを読み込み、計算結果を書き込むだけ
' テスト：VBA IDE での手動テスト（VBA Runnerでは不要）
Sub Main()
    Dim data As Variant
    data = Sheets("SalesData").Range("A1:A5").Value
    
    Dim result As Long
    result = CalculateTotalFromArray(data)
    
    Sheets("Report").Range("C1").Value = result
End Sub
```

**メリット**:
- `CalculateTotalFromArray()` は配列のみに依存 → 簡単にテスト可能
- 計算ロジックが明確 → テストコードがシンプル
- Excel mock は不要

---

## 原則2: 単体テストは Domain Logic のみ対象

### テスト可能な VBA コード構成

```
┌──────────────────────────────┐
│  Domain Logic 層             │  ← 単体テスト対象
│  (計算、データ処理など)      │     mock 不要
│  ・CalculateTotalFromArray   │     シンプルな単体テスト
│  ・FilterSalesData           │
│  ・FormatCurrency            │
└──────────────────────────────┘
           ▲
           │ 配列や値を渡す
           │
┌──────────────────────────────┐
│  Excel I/O 層                │  ← VBA IDE での手動テスト
│  (Sheets, Range操作)        │     VBA Runnerでテストしない
│  ・Main()                    │
│  ・LoadDataFromSheet         │
│  ・SaveResultToSheet         │
└──────────────────────────────┘
```

### テストコード例

```typescript
import { VBARunner } from './test-libs/test-runner';

describe('Sales Calculation', () => {
  const vbaRunner = new VBARunner('macro.bas');

  // ✅ 推奨：Domain Logic のみテスト
  it('should calculate total correctly', () => {
    const result = vbaRunner.run('CalculateTotalFromArray', [
      [100, 200, 300, 150, 250]
    ]);
    expect(result).toBe(1000);
  });

  // ✅ 推奨：境界値テスト
  it('should handle single value', () => {
    const result = vbaRunner.run('CalculateTotalFromArray', [[500]]);
    expect(result).toBe(500);
  });

  // ✅ 推奨：空配列テスト
  it('should return 0 for empty array', () => {
    const result = vbaRunner.run('CalculateTotalFromArray', [[]]);
    expect(result).toBe(0);
  });

  // ❌ 非推奨：Excel I/O まで含める（複雑）
  // it('should save result to Excel', () => {
  //   vbaRunner.setupMocks({
  //     Sheets: {
  //       'SalesData': { Range: { 'A1:A5': { Value: [100, 200, 300, 150, 250] } } },
  //       'Report': { Range: { 'C1': { Value: 0 } } }
  //     }
  //   });
  //   vbaRunner.run('Main', []);
  //   // ... 複雑な検証 ...
  // });
});
```

---

## 原則3: Excel I/O テストは VBA IDE で行う

### Excel I/O テストチェックリスト

**VBA IDE（Excel 内）で手動確認**:

```vba
' 手動テスト手順
' 1. test-fixtures/sales-data.xlsx を開く
' 2. VBA IDE で Main() を実行
' 3. Report.C1 に 1000 が入っているか確認
' 4. 予期した副作用が発生しているか確認
```

**チェックリスト**:
- [ ] データが正しく読み込まれた
- [ ] 計算結果が正しく書き込まれた
- [ ] 予期しない Excel 操作が発生していない
- [ ] パフォーマンスが許容範囲

### 自動テストが不要な理由

Excel I/O を VBA Runnerでテストすると：

```
テストコード ← mock 定義 → 本番コード
  ↑                         ↑
   ───複雑度が逆転───
```

- mock の複雑度 > 本番コードの複雑度
- テストの保守コストが上昇
- テストがテスト自体の正確性を検証する必要が出る（無限ループ）

**結論**: Excel I/O は VBA IDE での手動テストが最適

---

## パターン1: 純粋関数（テスト最適）

### VBA コード

```vba
' 入力：配列のみ
' 出力：計算結果
' 副作用：なし
Function SumArray(arr As Variant) As Long
    Dim i As Integer
    Dim total As Long
    For i = LBound(arr) To UBound(arr)
        total = total + arr(i)
    Next i
    SumArray = total
End Function

Function AverageArray(arr As Variant) As Double
    If UBound(arr) < LBound(arr) Then
        AverageArray = 0
        Exit Function
    End If
    AverageArray = SumArray(arr) / (UBound(arr) - LBound(arr) + 1)
End Function
```

### テストコード

```typescript
describe('Pure Functions', () => {
  const vbaRunner = new VBARunner('math.bas');

  it('SumArray: [100, 200, 300] = 600', () => {
    expect(vbaRunner.run('SumArray', [[100, 200, 300]])).toBe(600);
  });

  it('AverageArray: [100, 200, 300] = 200', () => {
    expect(vbaRunner.run('AverageArray', [[100, 200, 300]])).toBe(200);
  });

  it('AverageArray: empty = 0', () => {
    expect(vbaRunner.run('AverageArray', [[]])).toBe(0);
  });
});
```

**特徴**: シンプル、高速、確実

---

## パターン2: パラメータ化テスト（複数ケース）

### VBA コード

```vba
' ビジネスロジック：割引計算
Function CalculateDiscount(price As Double, discount As Double) As Double
    If discount < 0 Or discount > 1 Then
        Err.Raise 11  ' Division by zero (代わりにエラー)
    End If
    CalculateDiscount = price * (1 - discount)
End Function
```

### テストコード

```typescript
describe('Discount Calculation', () => {
  const vbaRunner = new VBARunner('pricing.bas');

  const testCases = [
    { price: 1000, discount: 0.1, expected: 900 },
    { price: 1000, discount: 0.5, expected: 500 },
    { price: 1000, discount: 0, expected: 1000 },
    { price: 1000, discount: 1, expected: 0 },
  ];

  testCases.forEach(({ price, discount, expected }) => {
    it(`Discount ${discount * 100}% on $${price} = $${expected}`, () => {
      const result = vbaRunner.run('CalculateDiscount', [price, discount]);
      expect(result).toBe(expected);
    });
  });

  it('should reject invalid discount', () => {
    expect(() => {
      vbaRunner.run('CalculateDiscount', [1000, -0.1]);
    }).toThrow();
  });
});
```

---

## パターン3: 状態変更を伴うロジック（Controller）

### VBA コード

```vba
' 状態を保持する（クラスまたはモジュール変数）
Private lastCalculatedValue As Long

Function GetLastValue() As Long
    GetLastValue = lastCalculatedValue
End Function

Function ProcessAndStore(value As Long) As Long
    ' ビジネスロジック
    lastCalculatedValue = value * 2
    ProcessAndStore = lastCalculatedValue
End Function

' Excel I/O（テストしない）
Sub UpdateSheet()
    Dim data As Variant
    data = Sheets("Input").Range("A1").Value
    Dim result As Long
    result = ProcessAndStore(data)
    Sheets("Output").Range("A1").Value = result
End Sub
```

### テストコード

```typescript
describe('Stateful Processing', () => {
  const vbaRunner = new VBARunner('stateful.bas');

  it('should process value and store state', () => {
    const result = vbaRunner.run('ProcessAndStore', [100]);
    expect(result).toBe(200);
    expect(vbaRunner.run('GetLastValue', [])).toBe(200);
  });

  it('should update state on each call', () => {
    vbaRunner.run('ProcessAndStore', [50]);
    vbaRunner.run('ProcessAndStore', [75]);
    expect(vbaRunner.run('GetLastValue', [])).toBe(150);
  });
});
```

---

## パターン4: エラーハンドリング

### VBA コード

```vba
Function SafeDivide(numerator As Long, denominator As Long) As Double
    If denominator = 0 Then
        Err.Raise 11  ' Division by zero
    End If
    SafeDivide = numerator / denominator
End Function

Function SafeDivideWithDefault(numerator As Long, denominator As Long, _
                                defaultValue As Double) As Double
    On Error GoTo ErrorHandler
    SafeDivideWithDefault = numerator / denominator
    Exit Function
ErrorHandler:
    SafeDivideWithDefault = defaultValue
End Function
```

### テストコード

```typescript
describe('Error Handling', () => {
  const vbaRunner = new VBARunner('safe-math.bas');

  it('should raise error on division by zero', () => {
    expect(() => {
      vbaRunner.run('SafeDivide', [10, 0]);
    }).toThrow();
  });

  it('should return default on error', () => {
    const result = vbaRunner.run('SafeDivideWithDefault', [10, 0, -1]);
    expect(result).toBe(-1);
  });

  it('should divide normally', () => {
    const result = vbaRunner.run('SafeDivide', [10, 2]);
    expect(result).toBe(5);
  });
});
```

---

## ベストプラクティス

### 1. VBA 関数は単一責任を持つ

```vba
' ❌ 悪い例（複数責任）
Function ProcessSalesData()
    Dim data As Variant
    data = Sheets("Raw").Range("A1:A100").Value          ' 責務1: 読込
    Dim filtered As Variant
    filtered = FilterOutliners(data)                      ' 責務2: フィルタリング
    Dim total As Long
    total = SumArray(filtered)                            ' 責務3: 計算
    Sheets("Report").Range("C1").Value = total            ' 責務4: 書込
End Function

' ✅ 良い例（関数分割）
Function FilterOutliners(data As Variant) As Variant
    ' フィルタリングロジックのみ
End Function

Function ProcessFilteredData(filteredData As Variant) As Long
    ' 計算ロジックのみ
End Function

Sub ProcessSalesData()
    ' I/O は Sub で（手動テスト対象）
    Dim data As Variant
    data = Sheets("Raw").Range("A1:A100").Value
    Dim filtered As Variant
    filtered = FilterOutliners(data)
    Dim total As Long
    total = ProcessFilteredData(filtered)
    Sheets("Report").Range("C1").Value = total
End Sub
```

### 2. 入力パラメータは具体的に

```vba
' ❌ 暗黙的（テストしづらい）
Function UpdateInventory()
    ' グローバル変数に依存
End Function

' ✅ 明示的（テスト可能）
Function CalculateNewInventory(currentStock As Long, soldUnits As Long, _
                                restockAmount As Long) As Long
    CalculateNewInventory = currentStock - soldUnits + restockAmount
End Function
```

> **vba-analyzer**: Excel オブジェクトへの直接アクセスを含む関数は `excelAccessCount` / `excelAccessSamples` として報告されます。パラメータなしで Excel に依存している関数はリファクタリング候補の筆頭です。

### 3. 関連パラメータは Type にまとめる

引数が増えてきた場合は、**ドメインとして意味のあるまとまり**を `Type` にまとめて渡すことを推奨します。
個々の引数を並べるよりも呼び出し側・テスト側の可読性が上がり、将来的なパラメータ追加にも強くなります。

> **注意**: 「同じ関数に渡すから」という理由だけでパラメータを1つの `Type` にまとめないでください。
> ドメインとして意味のあるまとまり（例：在庫状態、注文情報、期間指定）であることが前提です。
> 無関係なパラメータを1つの `Type` に詰め込むと、責務が曖昧になりテストや再利用がかえって困難になります。

```vba
' ❌ 引数が増えると呼び出しが煩雑になる
Function CalculateNewInventory(currentStock As Long, soldUnits As Long, _
                                restockAmount As Long, minStock As Long, _
                                maxStock As Long) As Long
    ' ...
End Function

' ✅ ドメインとしてまとまりのあるパラメータを Type にまとめる
Type InventoryParams
    CurrentStock  As Long
    SoldUnits     As Long
    RestockAmount As Long
    MinStock      As Long
    MaxStock      As Long
End Type

Function CalculateNewInventory(p As InventoryParams) As Long
    Dim newStock As Long
    newStock = p.CurrentStock - p.SoldUnits + p.RestockAmount
    If newStock < p.MinStock Then newStock = p.MinStock
    If newStock > p.MaxStock Then newStock = p.MaxStock
    CalculateNewInventory = newStock
End Function
```

TypeScript 側のテストでは、UDT は連想配列として渡します。

```typescript
const result = vbaRunner.run('CalculateNewInventory', [{
    CurrentStock: 100, SoldUnits: 30, RestockAmount: 50,
    MinStock: 10, MaxStock: 200
}]);
assert.strictEqual(result, 120);
```

> **vba-analyzer**: 変数名に共通の接頭辞（例: `inv_Stock`, `inv_Min`, `inv_Max`）が見られる場合、`prefixClusters` としてグルーピング候補を報告します。これは `Type` 化の手がかりになります。

同じ構造を複数のテストで使う場合は、TypeScript 側でも型を定義すると型安全になり可読性も上がります。

```typescript
interface InventoryParams {
    CurrentStock:  number;
    SoldUnits:     number;
    RestockAmount: number;
    MinStock:      number;
    MaxStock:      number;
}

function calcInventory(p: InventoryParams): number {
    return vbaRunner.run('CalculateNewInventory', [p]) as number;
}

assert.strictEqual(calcInventory({ CurrentStock: 100, SoldUnits: 30, RestockAmount: 50, MinStock: 10, MaxStock: 200 }), 120);
assert.strictEqual(calcInventory({ CurrentStock: 10,  SoldUnits: 20, RestockAmount: 5,  MinStock: 10, MaxStock: 200 }), 10); // MinStock に丸め
```

### 4. 振る舞いを持たせたい場合はクラスを検討する

> **vba-analyzer**: `prefixClusters` で検出された変数群に対して、代入・参照のパターンが複数のプロシージャにまたがる場合はクラス化の候補と考えられます（現時点では人手での判断が必要です）。

単純なデータの集約には `Type` が適切です。**バリデーションや複数の操作をデータと一緒に持たせたい**場合はクラスを検討します。

ただし VBA のクラスには制限があります（コンストラクタに引数を渡せない・継承がない・メソッドのオーバーロードがないなど）。振る舞いが単純なうちは `Type` + モジュール関数の組み合わせで十分なことが多く、クラス化は慎重に判断してください。

```vba
' Type + モジュール関数: 振る舞いが少なければこれで十分
Type InventoryParams
    CurrentStock  As Long
    SoldUnits     As Long
    RestockAmount As Long
    MinStock      As Long
    MaxStock      As Long
End Type

Function CalculateNewInventory(p As InventoryParams) As Long
    ' ...
End Function

' ✅ クラス: 状態保持・バリデーション・複数操作が必要になった場合
' InventoryManager.cls
'
' Private m_Stock As Long
' Private m_Min   As Long
' Private m_Max   As Long
'
' ' Class_Initialize は引数を取れないため Initialize メソッドで代用
' Sub Initialize(stock As Long, minVal As Long, maxVal As Long)
'     If minVal > maxVal Then Err.Raise 5, , "min > max"
'     m_Stock = stock : m_Min = minVal : m_Max = maxVal
' End Sub
'
' Function Apply(sold As Long, restock As Long) As Long
'     Dim n As Long
'     n = m_Stock - sold + restock
'     If n < m_Min Then n = m_Min
'     If n > m_Max Then n = m_Max
'     m_Stock = n   ' 状態を更新
'     Apply = n
' End Function
```

### 5. Excel オブジェクト依存は Sub に限定

> **vba-analyzer**: `excelMockTargets` に列挙された Function は Excel 依存を持つリファクタリング候補です。`Sheets`, `Range`, `Application` などのアクセスを含む Function を検出し、Sub への切り出し対象として報告します。
>
> ```bash
> node test-libs/vba-analyzer.cjs src/vba/ --json | jq '.excelMockTargets'
> ```

```vba
' ✅ 推奨パターン
Function CalculateXxx(param1 As ..., param2 As ...) As ...
    ' Domain Logic のみ
End Function

Sub ExecuteXxx()
    ' I/O のみ
    Dim input As Variant
    input = Sheets(...).Range(...).Value
    Dim output As ...
    output = CalculateXxx(input, ...)
    Sheets(...).Range(...).Value = output
End Sub
```

---

## テスト環境のセットアップ

### 最小限の準備

```typescript
import { VBARunner } from './test-libs/test-runner';

// テストファイル
const vbaRunner = new VBARunner('src/vba/business-logic.bas');

// テスト実行
describe('Business Logic Tests', () => {
  it('should calculate correctly', () => {
    const result = vbaRunner.run('MyFunction', [input]);
    expect(result).toBe(expected);
  });
});
```

### VBA IDE での補助テスト（手動）

```vba
' Debug パネルで実行
Sub DebugTest()
    ' Domain Logic のテスト（確認用）
    Debug.Print CalculateTotalFromArray(Array(100, 200, 300))  ' → 600
    
    ' Excel I/O のテスト（手動で Sheet を確認）
    Call Main()
    ' Report.C1 に値が入ったか目視確認
End Sub
```

---

## よくある質問

### Q1: Excel オブジェクトをテストしたい場合は？

**A**: VBA IDE で手動テストしてください。VBA Runnerでのテストは推奨しません。

理由：
- mock の複雑度が高い
- テストメンテナンスのコストが本番コードより高い
- 「Excel との動作仕様が正しいか」は VBA IDE でのみ検証可能

### Q2: 既存マクロが Excel I/O と計算が混在している場合は？

**A**: リファクタリングしてください。以下の手順：

1. Domain Logic を抽出（計算部分を Function に）
2. その Function をテスト
3. Sub は VBA IDE で手動テスト

### Q3: 複雑なデータ構造（オブジェクトのネスト）をテストする場合は？

**A**: 配列や辞書を活用：

```vba
' VBA: 複雑な構造
Type SalesRecord
    Month As String
    Amount As Long
    Category As String
End Type

Function ProcessRecords(records() As SalesRecord) As Long
    ' 計算
End Function

' TypeScript: テストで配列のオブジェクトを渡す
vbaRunner.run('ProcessRecords', [[
    { Month: 'Jan', Amount: 100, Category: 'A' },
    { Month: 'Feb', Amount: 200, Category: 'B' }
]]);
```

---

## まとめ

| テスト対象 | 手段 | 推奨度 |
|-----------|------|--------|
| Domain Logic（計算、フィルタリング等） | VBA Runner + 単体テスト | ✅ 推奨 |
| Excel I/O | VBA IDE + 手動テスト | ✅ 推奨 |
| Excel オブジェクト mock | VBA Runner | ❌ 非推奨 |
| 統合テスト | VBA IDE | ✅ 必要に応じて |

**黄金則**: 
> **Domain Logic をシンプルに、Excel I/O を最小限に。単体テストは Domain Logic のみ対象。**

