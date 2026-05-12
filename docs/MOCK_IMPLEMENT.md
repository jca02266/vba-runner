# 軽量 VBA モック実装ガイド

## はじめに

このドキュメントは、**Excel オブジェクトを完全には分離できない場合の、軽量なモック実装方法**を説明します。

### 前提：まず Domain Logic 分離を試みる

必ず先に以下を確認してください：

- **`docs/TESTING_STRATEGY.md`** — Domain Logic と I/O の分離原則
- **`docs/REFACTORING_GUIDE.md`** — リファクタリングの実践手法

**結論**: 9 割のケースは Domain Logic 分離で解決します。以下の場合にのみこのガイドを参照：

- Domain Logic と I/O の分離が困難・不可能
- Excel シート操作がビジネスロジック本体である
- マイグレーション中で段階的なテスト化が必要

---

## 原則：軽量モック設計

### 原則 1: 「完全互換」を目指さない

```vba
' ❌ 間違い：Excel と完全に同じ挙動を再現しようとする
' → モックが膨大になり、保守不可能

' ✅ 正解：テストに必要な最小限だけ実装
' → 100 行のモックで十分
```

### 原則 2: 動的定義可能にする

```typescript
// ❌ 悪い：事前にすべてのモックを準備
const mockWorksheet = { /* 500行 */ };

// ✅ 良い：テスト時に必要なものだけ定義
const mockWorksheet = {
  Range: (address) => ({ Value: 100 })
};
```

### 原則 3: 「スタブ」で十分（スパイ機能は不要）

スタブ = **値を返すだけ**  
スパイ = 呼び出しを記録（必要ならテスト内で確認）

```typescript
// ✅ シンプルなスタブ
const mockMsgBox = (message) => 1;  // 常に OK を返す

// ❌ 複雑なスパイ（初期段階では不要）
const mockMsgBox = jest.fn().mockReturnValue(1);
```

---

## 実装パターン

### パターン 1: JavaScript オブジェクト モック（最も軽量）

**使い方**:
```typescript
// テスト対象：Sheets オブジェクトに依存する VBA コード
const code = `
Function GetSalesTotal() As Long
    GetSalesTotal = Sheets("Sales").Range("B2").Value
End Function
`;

// モック定義
const mockSheets = {
  Sales: {
    Range: (address) => ({
      Value: 150000
    })
  }
};

// VBA コンパイラにモックを注入
const ev = new Evaluator();
ev.globalEnv.set('Sheets', mockSheets);

// テスト実行
const result = ev.callProcedure('GetSalesTotal', []);
expect(result).toBe(150000);
```

**メリット**:
- シンプル
- ボイラープレートが少ない
- テスト内で直接定義可能

**デメリット**:
- 複雑な操作には不向き
- アドレス解析（"A1:B10"）を手動で実装

---

### パターン 2: 軽量 Mock クラス

複数の Range を扱う場合：

```typescript
// シンプルな Mock クラス
class MockRange {
  private data: Map<string, any> = new Map();
  
  setValue(address: string, value: any) {
    this.data.set(address.toLowerCase(), value);
  }
  
  getValue(address: string): any {
    return this.data.get(address.toLowerCase()) ?? 0;
  }
}

class MockWorksheet {
  private name: string;
  private ranges: MockRange;
  
  constructor(name: string) {
    this.name = name;
    this.ranges = new MockRange();
  }
  
  Range(address: string) {
    return {
      Value: this.ranges.getValue(address),
      set Value(val: any) {
        this.ranges.setValue(address, val);
      }
    };
  }
}

class MockApplication {
  private sheets: Map<string, MockWorksheet> = new Map();
  
  Sheets(nameOrIndex: string | number) {
    const name = typeof nameOrIndex === 'string' 
      ? nameOrIndex 
      : `Sheet${nameOrIndex}`;
    
    if (!this.sheets.has(name)) {
      this.sheets.set(name, new MockWorksheet(name));
    }
    return this.sheets.get(name)!;
  }
}

// テストで使用
const mockApp = new MockApplication();
ev.globalEnv.set('Sheets', (name) => mockApp.Sheets(name));

// テストコード
mockApp.Sheets("Sales").Range("B2").Value = 150000;
const result = ev.callProcedure('GetSalesTotal', []);
expect(result).toBe(150000);
```

**メリット**:
- 複数 Range のサポート
- 読み書き可能
- 拡張可能

---

### パターン 3: ビルダー パターン（初期化が複雑な場合）

```typescript
class MockSheetBuilder {
  private sheets: Map<string, Map<string, any>> = new Map();
  
  // チェーン可能な API
  addSheet(name: string): this {
    this.sheets.set(name, new Map());
    return this;
  }
  
  setCellValue(sheetName: string, address: string, value: any): this {
    if (!this.sheets.has(sheetName)) {
      this.sheets.set(sheetName, new Map());
    }
    this.sheets.get(sheetName)!.set(address.toLowerCase(), value);
    return this;
  }
  
  build() {
    return (name: string) => ({
      Range: (address: string) => ({
        Value: this.sheets.get(name)?.get(address.toLowerCase()) ?? 0
      })
    });
  }
}

// テストで使用
const mockSheets = new MockSheetBuilder()
  .addSheet("Sales")
  .setCellValue("Sales", "B2", 150000)
  .setCellValue("Sales", "B3", 200000)
  .build();

ev.globalEnv.set('Sheets', mockSheets);
```

---

### パターン 4: JSON ベース設定

複雑なセットアップをテスト外で定義：

```typescript
// test-fixtures.ts
export const salesSheetMock = {
  Sales: {
    cells: {
      'B2': 150000,
      'B3': 200000,
      'B4': 175000
    }
  }
};

// test.spec.ts
import { salesSheetMock } from './test-fixtures';

const mockSheets = (name: string) => ({
  Range: (address: string) => ({
    Value: salesSheetMock[name]?.cells[address.toUpperCase()] ?? 0
  })
});

ev.globalEnv.set('Sheets', mockSheets);
```

---

## 実装例：複数パターンの組み合わせ

### シナリオ：売上データを読んで計算する VBA コード

```vba
Function AnalyzeSales() As String
    Dim total As Long
    Dim count As Long
    Dim i As Long
    
    ' シートから読む
    Dim data As Variant
    data = Sheets("Data").Range("A2:A10").Value
    
    ' 計算
    For i = 1 To UBound(data)
        If data(i, 1) > 100000 Then
            total = total + data(i, 1)
            count = count + 1
        End If
    Next i
    
    AnalyzeSales = total & " (" & count & "件)"
End Function
```

### テスト実装（軽量モック）

```typescript
describe('AnalyzeSales', () => {
  let ev: Evaluator;
  
  beforeEach(() => {
    const code = `...VBA コード...`;
    const tokens = new Lexer(code).tokenize();
    const ast = new Parser(tokens).parse();
    ev = new Evaluator(() => {});
    ev.evaluate(ast);
  });
  
  it('売上データを集計する', () => {
    // モックセットアップ
    const mockData = [
      [50000],
      [150000],    // ← 100000 超えた
      [80000],
      [200000]     // ← 100000 超えた
    ];
    
    const mockSheets = (name: string) => ({
      Range: (address: string) => ({
        Value: mockData
      })
    });
    
    ev.globalEnv.set('Sheets', mockSheets);
    
    // 実行
    const result = ev.callProcedure('AnalyzeSales', []);
    
    // 検証
    expect(result).toBe('350000 (2件)');
  });
});
```

---

## よくあるユースケース

### Use Case 1: シート値の読み取り

```typescript
// モック
const mockSheets = (name: string) => ({
  Range: (address: string) => ({
    Value: 100
  })
});

// VBA
Function GetValue() As Long
  GetValue = Sheets("Data").Range("A1").Value
End Function
```

### Use Case 2: シート値の書き込み + 読み取り

```typescript
class MockRange {
  constructor(private value: any) {}
  
  get Value() { return this.value; }
  set Value(v: any) { this.value = v; }
}

const mockSheets = (name: string) => ({
  Range: (address: string) => new MockRange(0)
});

// VBA
Sub SetAndRead()
  Sheets("Data").Range("A1").Value = 100
  MsgBox Sheets("Data").Range("A1").Value  ' 100
End Sub
```

### Use Case 3: Range 配列操作

```typescript
const mockSheets = (name: string) => ({
  Range: (address: string) => ({
    Value: [[1, 2], [3, 4]]  // 2x2 配列
  })
});

// VBA
Function SumRange() As Long
  Dim data As Variant
  data = Sheets("Data").Range("A1:B2").Value
  ' data(1,1)=1, data(1,2)=2, data(2,1)=3, data(2,2)=4
End Function
```

---

## チェックリスト：モック実装の判断

モック実装が必要か判定：

- [ ] Domain Logic を Function に分離できたか？
- [ ] I/O は Sub に集約できたか？
- [ ] テスト対象の Sub は 20 行以内か？
- [ ] Excel オブジェクトの使用箇所は 5 個以内か？

**すべて Yes** → mocker不要  
**No が 1 個以上** → 以下を検討：
1. まず REFACTORING_GUIDE を参照してリファクタリング
2. それでも分離不可能 → このガイドのモック実装

---

## アンチパターン：避けるべき実装

### ❌ アンチパターン 1: 完全互換を目指す

```typescript
// ❌ これはしない：膨大で保守不可能
class CompleteExcelMock {
  // 100+ lines...
  Worksheet: {
    Range: {
      Value: any,
      Interior: { Color: number },
      Font: { Bold: boolean, Size: number },
      // ... 50+ プロパティ
    }
  }
}
```

### ❌ アンチパターン 2: テスト固有の複雑ロジック

```typescript
// ❌ これはしない：テストコードがビジネスロジック化
const mockRange = {
  Value: data.filter(x => x > threshold).reduce((a, b) => a + b, 0)
};
```

**正解**: ビジネスロジックは VBA Function に。モックはデータ返却だけ。

### ❌ アンチパターン 3: グローバルモック状態

```typescript
// ❌ これはしない：テスト間で状態が共有
let globalMockState = {};

beforeEach(() => {
  // globalMockState に依存：テスト分離が破れる
});
```

**正解**: 各テストで独立したモック。

---

## 関連ドキュメント

- **`docs/TESTING_STRATEGY.md`** — テスト設計原則（まずこれを読む）
- **`docs/REFACTORING_GUIDE.md`** — リファクタリング手法（分離できない場合はここで工夫）
- **`docs/TEST_FRAMEWORK_GUIDE.md`** — JavaScript テストフレームワーク活用

---

## 結論

**モック実装は「最後の手段」**です：

```
Domain Logic 分離（REFACTORING_GUIDE）
    ↓ 分離不可能な場合のみ
軽量モック実装（このガイド）
    ↓ さらに複雑な場合
Excel 統合テスト（VBA IDE での手動テスト）
```

可能な限り Domain Logic を分離し、必要な部分だけ軽量モックで対応してください。
