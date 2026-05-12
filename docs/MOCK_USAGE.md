# 軽量モック実装：使用ガイド

## 概要

このドキュメントは、`src/compiler/mock/MockWorksheet.ts` に実装された軽量モックの**使い方**を説明します。

### 対象コード

- **実装**: `src/compiler/mock/MockWorksheet.ts`
- **テスト例**: `tests/spec/mock-usage.test.ts`

---

## クイックスタート

### 1. 簡単な読み書き

```typescript
import { MockApplication } from '../../src/compiler/mock/MockWorksheet';

// モックアプリケーション作成
const mockApp = new MockApplication();
const ws = mockApp.Sheets('TestSheet');

// セルに値を設定
ws.setCellValue('A1', 100);

// セルから値を読む
const value = ws.Range('A1').Value;  // 100
```

### 2. VBA コードとの統合

```typescript
import { Lexer } from '../../src/compiler/lexer';
import { Parser } from '../../src/compiler/parser';
import { Evaluator } from '../../src/compiler/evaluator';
import { MockApplication } from '../../src/compiler/mock/MockWorksheet';

// VBA コード
const vbaCode = `
  Function GetTotal() As Long
    GetTotal = Sheets("Data").Range("A1").Value
  End Function
`;

// コンパイル
const tokens = new Lexer(vbaCode).tokenize();
const ast = new Parser(tokens).parse();
const ev = new Evaluator(() => {});
ev.evaluate(ast);

// モック設定
const mockApp = new MockApplication();
mockApp.Sheets('Data').setCellValue('A1', 100);

// Sheets 関数にモックを注入
ev.globalEnv.set('Sheets', (name: string) => {
  const sheetObj = mockApp.Sheets(name);
  return {
    Range: (address: string) => sheetObj.Range(address),
  };
});

// テスト実行
const result = ev.callProcedure('GetTotal', []);
console.log(result);  // 100
```

---

## API リファレンス

### MockApplication

アプリケーションレベルのモック（Excel.Application に相当）。

#### メソッド

```typescript
// ワークシートを取得（存在しない場合は自動作成）
Sheets(nameOrIndex: string | number): MockWorksheet

// ワークシート一覧（デバッグ用）
listSheets(): string[]

// すべてをクリア
clear(): void
```

### MockWorksheet

ワークシートレベルのモック（Worksheet に相当）。

#### メソッド

```typescript
// セルに値を設定
// - 単一セル: 'A1'
// - 範囲: 'A1:C5'
// - 配列: [[1,2,3], [4,5,6]]
setCellValue(address: string, value: any): void

// セルから値を取得
getCellValue(address: string): any

// Range オブジェクトを取得
Range(address: string): MockRange

// 全セルをダンプ（デバッグ用）
dump(): Record<string, any>

// ワークシート名を取得
get Name(): string
```

### MockRange

セル/範囲を表すオブジェクト（Range に相当）。

#### プロパティ

```typescript
// セルの値を読み書き
Value: any  // get/set
```

---

## 使用例

### 例 1: 単純な読み書き

```typescript
const mockApp = new MockApplication();
const ws = mockApp.Sheets('Sheet1');

ws.setCellValue('A1', 100);
ws.setCellValue('B1', 200);

const val1 = ws.Range('A1').Value;  // 100
const val2 = ws.Range('B1').Value;  // 200
```

### 例 2: 範囲データ

```typescript
const ws = mockApp.Sheets('Sheet1');

// 2D 配列で範囲を設定
const data = [
  [1, 2, 3],
  [4, 5, 6],
  [7, 8, 9]
];
ws.setCellValue('A1:C3', data);

// 範囲を取得
const range = ws.Range('A1:C3').Value;
console.log(range[0][0]);  // 1
console.log(range[1][1]);  // 5
console.log(range[2][2]);  // 9
```

### 例 3: VBA コードでの読み取り

```typescript
const vbaCode = `
  Function ReadRange() As Long
    Dim data As Variant
    data = Sheets("Data").Range("A1:A5").Value
    
    Dim sum As Long
    Dim i As Long
    sum = 0
    For i = 1 To UBound(data)
      sum = sum + data(i, 1)
    Next i
    
    ReadRange = sum
  End Function
`;

// セットアップ
const tokens = new Lexer(vbaCode).tokenize();
const ast = new Parser(tokens).parse();
const ev = new Evaluator(() => {});
ev.evaluate(ast);

const mockApp = new MockApplication();
const ws = mockApp.Sheets('Data');
ws.setCellValue('A1:A5', [[10], [20], [30], [40], [50]]);

// Sheets をモックに置き換え
ev.globalEnv.set('Sheets', (name: string) => {
  const sheetObj = mockApp.Sheets(name);
  return {
    Range: (address: string) => sheetObj.Range(address),
  };
});

const result = ev.callProcedure('ReadRange', []);
console.log(result);  // 150
```

### 例 4: VBA コードでの書き込み

```typescript
const vbaCode = `
  Sub SetValue()
    Sheets("Output").Range("B2").Value = 999
  End Sub
`;

const tokens = new Lexer(vbaCode).tokenize();
const ast = new Parser(tokens).parse();
const ev = new Evaluator(() => {});
ev.evaluate(ast);

const mockApp = new MockApplication();
const ws = mockApp.Sheets('Output');

ev.globalEnv.set('Sheets', (name: string) => {
  const sheetObj = mockApp.Sheets(name);
  return {
    Range: (address: string) => sheetObj.Range(address),
  };
});

ev.callProcedure('SetValue', []);
const value = ws.getCellValue('B2');
console.log(value);  // 999
```

### 例 5: 複数シート

```typescript
const mockApp = new MockApplication();

// 複数シートを作成・設定
mockApp.Sheets('Sheet1').setCellValue('A1', 100);
mockApp.Sheets('Sheet2').setCellValue('A1', 200);
mockApp.Sheets('Sheet3').setCellValue('A1', 300);

// 各シートから読み取り
const v1 = mockApp.Sheets('Sheet1').Range('A1').Value;  // 100
const v2 = mockApp.Sheets('Sheet2').Range('A1').Value;  // 200
const v3 = mockApp.Sheets('Sheet3').Range('A1').Value;  // 300
```

---

## テスト実行方法

```bash
# モック実装のテストを実行
./node_modules/.bin/esbuild tests/spec/mock-usage.test.ts \
  --bundle --outfile=tests/spec/mock-usage.cjs --platform=node \
  && node tests/spec/mock-usage.cjs
```

出力例：
```
[Test 1] Simple Read/Write
[PASS] Simple Read/Write

[Test 2] Cell Write
[PASS] Cell Write

[Test 3] Range (Array)
[PASS] Range (Array)

[Test 4] VBA Code + Mock Integration
[PASS] VBA Code + Mock Integration

[Test 5] Multiple Sheets
[PASS] Multiple Sheets

[Test 6] Practical Example: Sales Calculation
[PASS] Practical Example: Sales Calculation

✅ Mock Implementation: 全テスト通過
```

---

## 対応アドレスフォーマット

モックが対応しているセルアドレス：

| フォーマット | 例 | 説明 |
|-------------|-----|------|
| 単一セル | `A1` | 1 つのセル |
| 範囲 | `A1:C5` | 矩形範囲 |
| 列指定 | `A:A` | 対応予定 |
| 行指定 | `1:10` | 対応予定 |

---

## 制限事項

現在の実装では以下は**未対応**です：

| 機能 | 状態 | 説明 |
|------|------|------|
| `Interior.Color` | ❌ | 書式設定 |
| `Font` | ❌ | フォント設定 |
| `Offset()` | ❌ | 相対位置参照 |
| `Select()` | ❌ | セル選択 |
| `Copy()` | ❌ | コピー操作 |
| `Paste()` | ❌ | ペースト操作 |
| `UsedRange` | ❌ | 使用範囲 |

**必要な場合**は、以下の方法を検討：

1. **ドメインロジック化** — Sheets 操作をドメイン関数に分離（推奨）
2. **モック拡張** — `MockWorksheet` に機能を追加
3. **手動テスト** — VBA IDE で Excel を使ってテスト

---

## よくある質問（FAQ）

### Q1: 書式設定（Color, Font）をモックしたい

**A**: MOCK_IMPLEMENT.md のパターン 2（Mock クラス拡張）を参照し、MockWorksheet を拡張してください。

```typescript
interface MockFormatting {
  Color?: number;
  Bold?: boolean;
  Size?: number;
}

class MockRangeWithFormat {
  Value: any;
  Interior: { Color?: number };
  Font: MockFormatting;
}
```

### Q2: 複数行を効率的に設定したい

**A**: 配列で一括設定：

```typescript
const data = [
  [1, 'A'],
  [2, 'B'],
  [3, 'C']
];
ws.setCellValue('A1:B3', data);
```

### Q3: Union Range（複数の離れた範囲）を使いたい

**A**: 現在非対応。範囲を分けて設定：

```typescript
ws.setCellValue('A1:A5', values1);
ws.setCellValue('C1:C5', values2);
```

### Q4: 名前付き範囲をサポートしたい

**A**: MockWorksheet に `namedRange` マッピングを追加できます。

```typescript
class MockWorksheet {
  private namedRanges: Map<string, string> = new Map();
  
  setNamedRange(name: string, address: string) {
    this.namedRanges.set(name, address);
  }
}
```

---

## 関連ドキュメント

- **`docs/MOCK_IMPLEMENT.md`** — モック実装の設計原則
- **`docs/REFACTORING_GUIDE.md`** — ドメインロジック分離（推奨）
- **`tests/spec/mock-usage.test.ts`** — テスト実装例
