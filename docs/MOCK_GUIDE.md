# VBA モック実装ガイド

## はじめに

このドキュメントは、**VBA コードが Excel オブジェクトに依存する場合のテスト方法**を、段階的に学びます。

### 前提条件

本ガイドを読む前に、以下を確認してください：

- **`docs/TESTING_STRATEGY.md`** — Domain Logic と I/O の分離原則
- **`docs/REFACTORING_GUIDE.md`** — リファクタリングの実践手法

**重要**: モックは「最後の手段」です。可能な限り Domain Logic を分離してください。

### ドキュメント構成

```
📍 Part 1: このプロジェクトの MockWorksheet を使う（すぐに使える）
  → 実装済みのモックを使ってテストを書く方法

📍 Part 2: MockWorksheet を参考に自分でモックを実装する（カスタマイズ）
  → MockWorksheet の設計を理解して独自モックを作成

📍 Part 3: モック設計の一般論（背景知識）
  → なぜこの設計なのか、いつモック化すべきか
```

**推奨される読み方**:
- 急いでいる → Part 1 のみ
- 拡張が必要 → Part 1 + 2
- 完全に理解したい → Part 1 + 2 + 3

---

# 📍 Part 1: このプロジェクトの MockWorksheet を使う

## 1. クイックスタート

### 最小限のコード例

```typescript
import { MockApplication } from '../../src/compiler/mock/MockWorksheet';
import { Lexer } from '../../src/compiler/lexer';
import { Parser } from '../../src/compiler/parser';
import { Evaluator } from '../../src/compiler/evaluator';

// VBA コードの定義
const vbaCode = `
  Function GetTotal() As Long
    GetTotal = Sheets("Data").Range("A1").Value
  End Function
`;

// VBA のコンパイル
const tokens = new Lexer(vbaCode).tokenize();
const ast = new Parser(tokens).parse();
const ev = new Evaluator(() => {});
ev.evaluate(ast);

// モック設定
const mockApp = new MockApplication();
mockApp.Sheets('Data').setCellValue('A1', 100);

// Sheets 関数をモックに置き換え
ev.getGlobalEnv().set('Sheets', (name: string) => {
  const sheetObj = mockApp.Sheets(name);
  return {
    Range: (address: string) => sheetObj.Range(address),
  };
});

// テスト実行
const result = ev.callProcedure('GetTotal', []);
console.log(result);  // 100
```

### 実行方法

```bash
./node_modules/.bin/esbuild your-test.ts \
  --bundle --outfile=your-test.cjs --platform=node \
  && node your-test.cjs
```

---

## 2. API リファレンス

### MockApplication

アプリケーションレベルのモック（Excel.Application に相当）。

```typescript
// ワークシートを取得（存在しない場合は自動作成）
Sheets(nameOrIndex: string | number): MockWorksheet

// ワークシート一覧を取得（デバッグ用）
listSheets(): string[]

// すべてをクリア
clear(): void
```

**例**:
```typescript
const mockApp = new MockApplication();
const ws1 = mockApp.Sheets('Sheet1');  // 名前で取得
const ws2 = mockApp.Sheets(1);         // インデックスで取得
```

### MockWorksheet

ワークシートレベルのモック（Worksheet に相当）。

```typescript
// セルに値を設定
// 単一セル: 'A1'
// 範囲: 'A1:C5'
// 配列: [[1,2,3], [4,5,6]]
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

**例**:
```typescript
const ws = mockApp.Sheets('Data');

// 単一セル
ws.setCellValue('A1', 100);
ws.setCellValue('B1', 'Alice');

// 範囲
ws.setCellValue('A1:B3', [
  [1, 'Alice'],
  [2, 'Bob'],
  [3, 'Charlie']
]);

// 読み取り
const val = ws.getCellValue('A1');
const range = ws.Range('A1:B3');
```

### MockRange

セル/範囲を表すオブジェクト（Range に相当）。

```typescript
// セルの値を読み書き
Value: any  // get/set
```

**例**:
```typescript
const range = ws.Range('A1');
console.log(range.Value);     // 読み取り
range.Value = 200;             // 書き込み
```

---

## 3. 使用例（5 段階）

### 例 1: 単純な読み書き

```typescript
const mockApp = new MockApplication();
const ws = mockApp.Sheets('Sheet1');

ws.setCellValue('A1', 100);
ws.setCellValue('B1', 200);

const val1 = ws.Range('A1').Value;  // 100
const val2 = ws.Range('B1').Value;  // 200

console.log('[PASS] Simple Read/Write');
```

### 例 2: セル書き込み

```typescript
const mockApp = new MockApplication();
const ws = mockApp.Sheets('TestSheet');

ws.Range('A1').Value = 10;
const original = ws.Range('A1').Value;
console.log(original);  // 10

ws.Range('A1').Value = 50;
const updated = ws.Range('A1').Value;
console.log(updated);  // 50

console.log('[PASS] Cell Write');
```

### 例 3: 範囲操作（配列）

```typescript
const mockApp = new MockApplication();
const ws = mockApp.Sheets('TestSheet');

const data = [[1, 2, 3], [4, 5, 6], [7, 8, 9]];
ws.setCellValue('A1:C3', data);

const range = ws.Range('A1:C3').Value;
console.log(range[0][0]);  // 1
console.log(range[1][1]);  // 5
console.log(range[2][2]);  // 9

console.log('[PASS] Range (Array)');
```

### 例 4: VBA コード + モック統合

```typescript
const vbaCode = `
  Function SumRange() As Long
    Dim data As Variant
    data = Sheets("Data").Range("A1:A5").Value
    Dim sum As Long
    Dim i As Long
    sum = 0
    For i = 1 To UBound(data)
        sum = sum + data(i, 1)
    Next i
    SumRange = sum
  End Function
`;

const tokens = new Lexer(vbaCode).tokenize();
const ast = new Parser(tokens).parse();
const ev = new Evaluator(() => {});
ev.evaluate(ast);

const mockApp = new MockApplication();
const ws = mockApp.Sheets('Data');
ws.setCellValue('A1:A5', [[10], [20], [30], [40], [50]]);

ev.getGlobalEnv().set('Sheets', (name: string) => {
  const sheetObj = mockApp.Sheets(name);
  return {
    Range: (address: string) => sheetObj.Range(address),
  };
});

const result = ev.callProcedure('SumRange', []);
console.log(result);  // 150

console.log('[PASS] VBA Code + Mock Integration');
```

### 例 5: 複数シート管理

```typescript
const mockApp = new MockApplication();

mockApp.Sheets('Sheet1').setCellValue('A1', 100);
mockApp.Sheets('Sheet2').setCellValue('A1', 200);
mockApp.Sheets('Sheet3').setCellValue('A1', 300);

const v1 = mockApp.Sheets('Sheet1').Range('A1').Value;  // 100
const v2 = mockApp.Sheets('Sheet2').Range('A1').Value;  // 200
const v3 = mockApp.Sheets('Sheet3').Range('A1').Value;  // 300

console.log('[PASS] Multiple Sheets');
```

---

## 4. テスト実行方法

### esbuild でコンパイル・実行

```bash
# テストファイルをコンパイル
./node_modules/.bin/esbuild tests/spec/your-test.ts \
  --bundle --outfile=tests/spec/your-test.cjs --platform=node

# 実行
node tests/spec/your-test.cjs
```

### 出力例

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

✅ Mock Implementation: 全テスト通過
```

### よくあるエラーと対応

| エラー | 原因 | 対応 |
|--------|------|------|
| `Cannot find module 'Sheets'` | `Sheets` 関数をモックに置き換えてない | `ev.getGlobalEnv().set('Sheets', ...)` を実行 |
| `undefined is not an object` | Range オブジェクトが正しく返されていない | `Range()` が MockRange を返しているか確認 |
| `ReferenceError: Sheets is not defined` | モックが注入されていない | `getGlobalEnv()` を使用しているか確認 |

---

## 5. 対応機能と制限事項

### 対応アドレスフォーマット

| フォーマット | 例 | 説明 |
|-------------|-----|------|
| 単一セル | `A1`, `Z100` | 1 つのセル |
| 範囲 | `A1:C5`, `B2:D10` | 矩形範囲 |
| 列指定 | `A:A` | 対応予定 |
| 行指定 | `1:10` | 対応予定 |

### 現在対応していない機能

| 機能 | 状態 | 説明 |
|------|------|------|
| `Interior.Color` | ❌ | 書式設定（セルの背景色） |
| `Font` | ❌ | フォント設定 |
| `Offset()` | ❌ | 相対位置参照 |
| `Select()` | ❌ | セル選択 |
| `Copy()` | ❌ | コピー操作 |
| `Paste()` | ❌ | ペースト操作 |
| `UsedRange` | ❌ | 使用範囲 |

### ⚠️ MockApplication の制限: Application プロパティは「無視」される

`MockApplication` が実装しているのは `Sheets()` メソッドのみ。
`Application.ScreenUpdating`・`Application.Calculation`・`Application.EnableEvents` などのプロパティは**実装されていない**。

しかしエラーにもならない。理由は evaluator の動作にある:

```
VBA:  Application.ScreenUpdating = False
        ↓
evaluator: obj['screenupdating'] = false   ← JS オブジェクトへの動的プロパティ代入
```

JavaScript オブジェクトは未定義プロパティへの代入を無視しないため、エラーなくプロパティが生える。
読み取り時（`state = Application.ScreenUpdating`）も同様に `undefined` が返るだけ。

**結果として何が起きるか:**

| コード | 実際の動作 |
|--------|-----------|
| `screenUpdateState = Application.ScreenUpdating` | `undefined` が代入される |
| `Application.ScreenUpdating = False` | mockApp に `screenupdating: false` が動的に生える（副作用なし） |
| `Application.ScreenUpdating = screenUpdateState` | `undefined` を書き戻す（副作用なし） |

**何がテストされていないか:**
- `ScreenUpdating` の保存・復元ロジックそのものは検証されない
- Excelのパフォーマンス設定（描画停止・手動計算）がテスト中に実際に変わることはない

**運用上の意味:**
スケジューリングやデータ変換のロジックをテストする目的では問題ない。
ただし「`Application.ScreenUpdating` の保存/復元が正しいこと」を検証したい場合は、
`Application` オブジェクト自体をプロパティを記録するモックに差し替える必要がある。

### 「こういう機能が欲しい」場合の対応方法

1. **Domain Logic で回避**（推奨）
   - 書式設定に依存したロジックは避ける
   - 色判定は値で判定（例：色コード → 文字列フラグ）

2. **MockWorksheet を拡張**（Part 2 参照）
   - `Interior`, `Font` などのプロパティを追加
   - `Offset()` メソッドを実装

3. **手動テスト**（VBA IDE で実施）
   - 最終確認は実 Excel で検証
   - 自動テストでカバーできない部分

---

# 📍 Part 2: MockWorksheet を参考に自分でモックを実装する

## 6. MockWorksheet.ts の実装を読み解く

### クラス構成

`src/compiler/mock/MockWorksheet.ts` の構造：

```
MockApplication（Excel.Application 相当）
  │
  └─ Sheets(nameOrIndex) → Map<string, MockWorksheet>
       │
       └─ MockWorksheet（Worksheet 相当）
            │
            ├─ cells: Map<string, any>（セル値の保存）
            └─ Range(address) → MockRange
                 │
                 └─ MockRange（Range 相当）
                      │
                      └─ Value: any（読み書き可能）
```

### アドレス解析ロジック

MockWorksheet は、アドレス文字列（"A1" や "A1:C5"）を解析して、セル座標に変換：

```typescript
private parseAddress(address: string): { row: number, col: number } | null {
  // "A1" → { row: 1, col: 1 }
  // "Z100" → { row: 100, col: 26 }
}

private parseRange(address: string): { start: {row, col}, end: {row, col} } | null {
  // "A1:C5" → { start: {row:1, col:1}, end: {row:5, col:3} }
}
```

### 2D 配列の扱い

範囲を 2D 配列で設定する場合：

```typescript
ws.setCellValue('A1:B3', [
  [1, 2],
  [3, 4],
  [5, 6]
]);
// 内部的には、各セルに展開されて保存される
// A1=1, B1=2, A2=3, B2=4, ...
```

読み取る場合は、範囲内のセルを 2D 配列として再構築：

```typescript
const range = ws.Range('A1:B3').Value;
// [[1, 2], [3, 4], [5, 6]] として返される
```

### 実装の特徴

1. **セル単位で保存** — 内部は `Map<"A1", value>` で単一セル単位
2. **アドレス正規化** — `"a1"`, `"A1"` どちらでも同じセルを参照
3. **動的 2D 配列変換** — 範囲読み取り時に動的に 2D 配列に変換
4. **複数シート対応** — MockApplication が複数の MockWorksheet を管理

---

## 7. 自作モックのパターン（4 つ）

### パターン 1: JavaScript オブジェクト（最軽量）

**最もシンプル**。テストが単純な場合に最適。

```typescript
const mockSheets = {
  Sales: {
    Range: (address) => ({
      Value: 150000
    })
  }
};

ev.getGlobalEnv().set('Sheets', (name) => mockSheets[name]);
```

**メリット**:
- コードが短い
- 依存関係がない
- セットアップが速い

**デメリット**:
- 複数セルには不向き
- アドレス解析を手動で実装

**使う場合**:
- 単一の値を返すだけ
- テストが 1-2 個のセルだけ

---

### パターン 2: Mock クラス（中程度）

**MockWorksheet を参考に設計**。複数セルの操作が必要な場合。

```typescript
class MockRange {
  constructor(private value: any) {}
  
  get Value() { return this.value; }
  set Value(v: any) { this.value = v; }
}

class MockWorksheet {
  private cells: Map<string, any> = new Map();
  
  Range(address: string) {
    const key = address.toUpperCase();
    if (!this.cells.has(key)) {
      this.cells.set(key, undefined);
    }
    const value = this.cells.get(key);
    return {
      get Value() { return value; },
      set Value(v: any) { this.cells.set(key, v); }
    };
  }
}

class MockApplication {
  private sheets: Map<string, MockWorksheet> = new Map();
  
  Sheets(name: string) {
    if (!this.sheets.has(name)) {
      this.sheets.set(name, new MockWorksheet());
    }
    return this.sheets.get(name)!;
  }
}

// テストで使用
const mockApp = new MockApplication();
ev.getGlobalEnv().set('Sheets', (name) => mockApp.Sheets(name));
```

**メリット**:
- 複数シート対応
- 読み書き可能
- 拡張が容易

**デメリット**:
- コードが増える
- セットアップが複雑

**使う場合**:
- 複数のシート・セルを扱う
- 読み書き両方が必要

---

### パターン 3: ビルダー パターン（初期化が複雑）

初期化データが多い場合、チェーン API で簡潔に書く。

```typescript
class MockSheetBuilder {
  private sheets: Map<string, Map<string, any>> = new Map();
  
  addSheet(name: string): this {
    this.sheets.set(name, new Map());
    return this;
  }
  
  setCellValue(sheetName: string, address: string, value: any): this {
    if (!this.sheets.has(sheetName)) {
      this.sheets.set(sheetName, new Map());
    }
    this.sheets.get(sheetName)!.set(address.toUpperCase(), value);
    return this;
  }
  
  build() {
    return (name: string) => ({
      Range: (address: string) => ({
        Value: this.sheets.get(name)?.get(address.toUpperCase()) ?? 0
      })
    });
  }
}

// テストで使用（初期化がきれい）
const mockSheets = new MockSheetBuilder()
  .addSheet("Sales")
  .setCellValue("Sales", "B2", 150000)
  .setCellValue("Sales", "B3", 200000)
  .setCellValue("Sales", "B4", 175000)
  .build();

ev.getGlobalEnv().set('Sheets', mockSheets);
```

**メリット**:
- 初期化コードが読みやすい
- チェーン API で簡潔
- テストの可読性向上

**デメリット**:
- ビルダークラスの実装が必要
- 一度ビルドするとデータ固定

**使う場合**:
- 複雑な初期化データ
- 複数テストで同じデータを使う

---

### パターン 4: JSON ベース設定（複数テストで共有）

テストデータを JSON で定義して、複数テストで共有。

```typescript
// test-fixtures.json
export const salesDataFixture = {
  Sales: {
    B2: 150000,
    B3: 200000,
    B4: 175000,
    B5: 180000
  },
  Inventory: {
    A1: 100,
    A2: 50,
    A3: 75
  }
};

// test.spec.ts
import { salesDataFixture } from './test-fixtures';

const mockSheets = (sheetName: string) => ({
  Range: (address: string) => ({
    Value: salesDataFixture[sheetName]?.[address.toUpperCase()] ?? 0
  })
});

ev.getGlobalEnv().set('Sheets', mockSheets);
```

**メリット**:
- テストデータを一元管理
- 複数テストで再利用
- テスト実装が簡潔

**デメリット**:
- 別ファイルを管理する必要
- 大規模データは管理困難

**使う場合**:
- 複数テストで同じデータ
- テストシナリオが標準化されている

---

## 8. 実装例：カスタムモックの作成

### 例 1: Interior.Color（セルの背景色）を追加

```typescript
class MockRangeWithFormat {
  constructor(
    private cellValue: any,
    private cellFormat: { color?: number } = {}
  ) {}
  
  get Value() { return this.cellValue; }
  set Value(v: any) { this.cellValue = v; }
  
  get Interior() {
    return {
      Color: this.cellFormat.color,
      set Color(c: number) {
        this.cellFormat.color = c;
      }
    };
  }
}

class MockWorksheetWithFormat {
  private cells: Map<string, MockRangeWithFormat> = new Map();
  
  Range(address: string) {
    const key = address.toUpperCase();
    if (!this.cells.has(key)) {
      this.cells.set(key, new MockRangeWithFormat(undefined));
    }
    return this.cells.get(key)!;
  }
}

// テスト
const ws = new MockWorksheetWithFormat();
ws.Range('A1').Value = 'Important';
ws.Range('A1').Interior.Color = 0xFF0000;  // 赤

console.log(ws.Range('A1').Value);          // 'Important'
console.log(ws.Range('A1').Interior.Color); // 0xFF0000
```

### 例 2: Offset()（相対位置参照）を追加

```typescript
class MockRangeWithOffset {
  constructor(
    private address: string,
    private worksheet: MockWorksheet
  ) {}
  
  get Value() { return this.worksheet.getCellValue(this.address); }
  set Value(v: any) { this.worksheet.setCellValue(this.address, v); }
  
  Offset(rowOffset: number, colOffset: number) {
    // "A1" → { row: 1, col: 1 }
    const [colLetter, rowNum] = this.address.match(/([A-Z]+)(\d+)/)!.slice(1);
    const newRow = parseInt(rowNum) + rowOffset;
    const newCol = String.fromCharCode(colLetter.charCodeAt(0) + colOffset);
    const newAddress = newCol + newRow;
    
    return new MockRangeWithOffset(newAddress, this.worksheet);
  }
}

// テスト
const ws = mockApp.Sheets('Data');
ws.setCellValue('A1', 100);
ws.setCellValue('A2', 200);

const range = ws.Range('A1') as MockRangeWithOffset;
const offsetRange = range.Offset(1, 0);  // A1 + 1行下 = A2
console.log(offsetRange.Value);  // 200
```

### 例 3: 名前付き範囲をサポート

```typescript
class MockWorksheetWithNamedRange {
  private cells: Map<string, any> = new Map();
  private namedRanges: Map<string, string> = new Map();
  
  setNamedRange(name: string, address: string) {
    this.namedRanges.set(name.toUpperCase(), address.toUpperCase());
  }
  
  Range(addressOrName: string) {
    let actualAddress = addressOrName.toUpperCase();
    
    // 名前付き範囲か判定
    if (this.namedRanges.has(actualAddress)) {
      actualAddress = this.namedRanges.get(actualAddress)!;
    }
    
    return {
      get Value() { return this.cells.get(actualAddress); },
      set Value(v: any) { this.cells.set(actualAddress, v); }
    };
  }
}

// テスト
const ws = new MockWorksheetWithNamedRange();
ws.setNamedRange('MyRange', 'A1:B5');
ws.Range('MyRange').Value = 100;
console.log(ws.Range('A1').Value);  // 100
```

---

# 📍 Part 3: モック設計の一般論

## 9. モック設計の原則

### 原則 1: 「完全互換」を目指さない

**❌ 間違い：Excel と完全に同じ挙動を再現**

```typescript
// こんなことはしない：膨大で保守不可能
class CompleteExcelMock {
  // 100+ lines...
  Worksheet: {
    Range: {
      Value: any,
      Interior: { Color: number },
      Font: { Bold: boolean, Size: number, Italic: boolean, Name: string },
      Offset: (r: number, c: number) => Range,
      Copy: () => void,
      Paste: () => void,
      PasteSpecial: (format: number) => void,
      // ... 50+ プロパティ
    }
  }
}
```

**✅ 正解：テストに必要な最小限だけ実装**

```typescript
// 最小限で十分（10-50行で十分）
class LightweightMock {
  Range(address: string) {
    return { Value: 0 };
  }
}
```

**理由**:
- モックが膨大になると、モック自体がバグの源
- 「何のためのテスト？」という本来目的を見失う
- 保守コストが増加

**判定基準**:
- 「テストに必要か？」を問う
- 「Excel との互換性」ではなく「テスト成功」が目的

---

### 原則 2: 動的定義可能にする

**❌ 悪い：事前にすべてのモックを準備**

```typescript
// セットアップが固い
const mockWorksheet = new MockWorksheet();
mockWorksheet.setCellValue('A1', 100);
mockWorksheet.setCellValue('A2', 200);
mockWorksheet.setCellValue('A3', 300);
// ... 100行

// 別のテストではコピー＆ペースト
const mockWorksheet2 = new MockWorksheet();
mockWorksheet2.setCellValue('A1', 500);
// ...
```

**✅ 良い：テスト時に必要なものだけ定義**

```typescript
// 軽量で柔軟
const mockSheets = (name: string) => ({
  Range: (address: string) => ({ Value: 100 })
});

// または
const mockApp = new MockApplication();
// 必要に応じて追加
mockApp.Sheets('Data').setCellValue('A1', 100);
```

**利点**:
- テストごとに異なるデータを簡単に用意
- モック定義がテストの近くにある（可読性向上）
- DRY 原則に従う

---

### 原則 3: 「スタブ」で十分（スパイ機能は不要）

**スタブ** = 値を返すだけ（受動的）  
**スパイ** = 呼び出しを記録（能動的）

**❌ スパイ（初期段階では不要）**

```typescript
// こういった複雑なスパイ機能は不要
const mockMsgBox = jest.fn().mockReturnValue(1);
mockMsgBox('Hello');
expect(mockMsgBox).toHaveBeenCalledWith('Hello');
expect(mockMsgBox).toHaveBeenCalledTimes(1);
```

**✅ スタブ（シンプルで十分）**

```typescript
// テストに必要な値を返すだけ
const mockMsgBox = (message: string) => 1;
const result = mockMsgBox('Hello');
expect(result).toBe(1);
```

**理由**:
- 「呼び出し回数の検証」は、テストが複雑すぎることの兆候
- Domain Logic が分離できていないことを示唆
- テストの目的（戻り値の正確性）に絞る

---

## 10. アンチパターン：避けるべき実装

### アンチパターン 1: テスト固有の複雑ロジック

**❌ これはしない：テストコードがビジネスロジック化**

```typescript
// モックが計算を行っている（テストの責務超越）
const mockRange = {
  Value: data.filter(x => x > threshold)
         .map(x => x * 2)
         .reduce((a, b) => a + b, 0)
};
```

**問題**:
- テストのモックが複雑になる
- モックのバグをテストしている状態に
- テスト実装者が計算ロジックを理解する必要

**✅ 正解**:

```typescript
// ビジネスロジックは VBA Function に
// モックはデータ返却だけ
const mockRange = { Value: 500 };  // 計算済みの値を返すだけ

// 計算は VBA で検証
const result = ev.callProcedure('CalculateSum', [100, 200, 200]);
expect(result).toBe(500);
```

---

### アンチパターン 2: グローバルモック状態

**❌ これはしない：テスト間で状態が共有**

```typescript
// グローバル状態：テスト分離が破れる
let globalMockState = { value: 0 };

it('test 1', () => {
  globalMockState.value = 100;
  // ...
});

it('test 2', () => {
  // test 1 の副作用に依存（テスト順序に依存）
  expect(globalMockState.value).toBe(100);
});
```

**問題**:
- テスト実行順序に依存
- テスト分離の原則を破る
- デバッグが困難

**✅ 正解**:

```typescript
it('test 1', () => {
  const mockApp = new MockApplication();  // 新しいインスタンス
  mockApp.Sheets('Data').setCellValue('A1', 100);
  // ...
});

it('test 2', () => {
  const mockApp = new MockApplication();  // 独立したインスタンス
  mockApp.Sheets('Data').setCellValue('A1', 200);
  // ...
});
```

---

### アンチパターン 3: モック実装に時間を使い過ぎ

**❌ これはしない：モック実装が本体より長い**

```typescript
// VBA コード: 50行
// モック実装: 500行  ← 危険信号！
```

**兆候**:
- モック実装が VBA コードより長い
- 複数の特殊ケースを処理している
-「テストが複雑すぎて理解できない」

**✅ 正解**:

```
モック実装 < VBA コード  が原則

長い場合は：
1. ビジネスロジックを分離（REFACTORING_GUIDE.md）
2. よりシンプルなモックパターンを採用
```

---

## 11. 判断フロー：いつモック化すべきか

### チェックリスト

モック実装が必要か判定：

- [ ] **Domain Logic を Function に分離できたか？**
  - 「はい」→ 以下に進む
  - 「いいえ」→ REFACTORING_GUIDE.md を参照してリファクタリング

- [ ] **I/O は Sub に集約できたか？**
  - 「はい」→ 以下に進む
  - 「いいえ」→ REFACTORING_GUIDE.md を参照

- [ ] **テスト対象の Sub は 20 行以内か？**
  - 「はい」→ 以下に進む
  - 「いいえ」→ さらに分割を検討

- [ ] **Excel オブジェクトの使用箇所は 5 個以内か？**
  - 「はい」→ モック実装を開始
  - 「いいえ」→ さらなるリファクタリング

### 判定結果

| 結果 | 対応 |
|------|------|
| **すべて Yes** | モック実装開始（このガイド Part 1） |
| **No が 1-2 個** | リファクタリング（REFACTORING_GUIDE.md） |
| **No が 3 個以上** | 大幅なリファクタリングが必要 |

---

## 12. テスト戦略：段階的なテスト化

```
┌─────────────────────────────────────┐
│ Unit Tests（単体テスト）             │
├─────────────────────────────────────┤
│ Domain Logic Functions のみ           │
│ → Excel 依存なし                     │
│ → 100% カバー可能                    │
│ テスト数：多い（関数ごと）            │
└─────────────────────────────────────┘
           ↓
┌─────────────────────────────────────┐
│ Integration Tests（統合テスト）      │
├─────────────────────────────────────┤
│ 複数関数 + I/O Sub                   │
│ → Mock Worksheet を使用              │
│ → Excel なしで実行                   │
│ テスト数：中程度（シナリオごと）     │
└─────────────────────────────────────┘
           ↓
┌─────────────────────────────────────┐
│ E2E Tests（エンドツーエンド）        │
├─────────────────────────────────────┤
│ 実際の Excel ファイルで実行           │
│ → VBA IDE で手動実行                 │
│ → UI 操作の最終確認                  │
│ テスト数：少ない（重要ケースのみ）   │
└─────────────────────────────────────┘
```

**推奨される配分**:
- Unit Tests: 70-80%（大部分）
- Integration Tests: 15-25%（主要シナリオ）
- E2E Tests: 5-10%（最終確認）

---

## 13. よくある質問（FAQ）

### Q1: MockWorksheet に機能を追加したい

**A**: Part 2 の「実装例：カスタムモックの作成」を参照してください。

例えば、`Interior.Color` を追加したい場合：

```typescript
class MockRangeWithFormat {
  constructor(private value: any) {}
  
  get Interior() {
    return { Color: 0xFF0000 };
  }
}
```

拡張ポイント：
- `Interior`, `Font` などのプロパティ追加
- `Offset()`, `Select()` などのメソッド追加
- `UsedRange` などの特殊な Range を実装

---

### Q2: 複数のテストで同じモック状態を使いたい

**A**: パターン 4（JSON ベース設定）を使用してください。

```typescript
// test-fixtures.ts
export const commonTestData = {
  Sales: { A1: 100, A2: 200 },
  Inventory: { A1: 50, A2: 75 }
};

// test1.spec.ts
import { commonTestData } from './test-fixtures';
const mockSheets = (name) => ({
  Range: (addr) => ({ Value: commonTestData[name][addr] })
});

// test2.spec.ts も同じ mockSheets を使用
```

---

### Q3: VBA コードのデバッグが難しい

**A**: 以下の方法を組み合わせてください：

1. **小さな Function ごとにテスト**
   ```typescript
   const result = ev.callProcedure('CalcLevel', [1, 3, 0.3]);
   console.log(result);
   ```

2. **モックデータを print して確認**
   ```typescript
   console.log('[DEBUG] mockApp state:', mockApp.dump());
   ```

3. **最終的には実 Excel で確認**
   - VBA IDE でステップ実行
   - Excel シートに手動でデータ入力

---

### Q4: 書式設定（Color, Font）が必要な場合は？

**A**: 3 つの対応方法：

1. **Domain Logic で回避（推奨）**
   - 色判定の代わりに値で判定
   - 例：赤 → 負の値, 緑 → 正の値

2. **MockWorksheet を拡張（Part 2 参照）**
   - `Interior.Color` を実装
   - テスト可能な形に

3. **手動テスト（VBA IDE）**
   - 自動テストでカバーできない部分
   - 実 Excel で最終確認

---

### Q5: Union Range（複数の離れた範囲）を使いたい

**A**: 現在の MockWorksheet は未対応。以下の対応方法：

**方法 1: 範囲を分けて設定**
```typescript
ws.setCellValue('A1:A5', values1);
ws.setCellValue('C1:C5', values2);
```

**方法 2: MockWorksheet を拡張**
```typescript
class MockWorksheetWithUnion {
  UnionRange(range1: string, range2: string) {
    return {
      Value: [
        this.Range(range1).Value,
        this.Range(range2).Value
      ]
    };
  }
}
```

**方法 3: ビジネスロジック化**
- Union Range の必要性を検討
- 複数の処理として分割

---

## 14. 関連ドキュメント

- **`docs/TESTING_STRATEGY.md`** — テスト設計の原則（最初に読む）
- **`docs/REFACTORING_GUIDE.md`** — Domain Logic 分離（モック前に読む）
- **`docs/INTEGRATION_TEST_EXAMPLE.md`** — 統合テストの実装例
- **`docs/TEST_FRAMEWORK_GUIDE.md`** — JavaScript テストフレームワーク活用
- **`tests/spec/mock-usage.test.ts`** — このプロジェクトのテスト実装例
- **`src/compiler/mock/MockWorksheet.ts`** — MockWorksheet の実装ソース

---

## まとめ

### モック実装の原則

```
Domain Logic 分離（最優先）
    ↓ 分離不可能な場合のみ
軽量モック実装（このガイド）
    ↓ さらに複雑な場合
Excel 統合テスト（VBA IDE での手動テスト）
```

### 重要な心がけ

1. **モックは最後の手段** — 可能な限り分離を試みる
2. **軽量を心がける** — 50 行のモックで十分
3. **テスト責務を混同しない** — モックは値を返すだけ
4. **テスト分離を守る** — 各テストは独立したモック

これらの原則に従えば、効果的で保守性の高いテストが実現できます。
