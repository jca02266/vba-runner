# VBA テスト：JS テストフレームワーク活用ガイド

## はじめに

このドキュメントは、**既存の JavaScript テストフレームワーク（Jest など）を活用して、VBA マクロの単体テストを効率的に書く方法**を説明します。

`docs/TESTING_STRATEGY.md` で説明した Domain Logic テストを、実装レベルで具体化するガイドです。

---

## 推奨構成

```
VBA コンパイラ + VBATest クラス
           ↓
JavaScript テストフレームワーク（Jest）
           ↓
Domain Logic の単体テスト（配列等）
```

**メリット**:
- 既存の Jest エコシステムを活用（すぐに使える）
- VBA コンパイラと TypeScript の統合
- 複雑な mock 不要（配列のみでテスト）

---

## 前提知識

**必ず先に読む**:
- `docs/TESTING_STRATEGY.md` - VBA テスト設計原則

本ドキュメントは、その原則を JavaScript 実装レベルで解説します。

---

## セットアップ

### 既存環境（推奨）

このプロジェクトは既に以下を備えています：

```bash
# TypeScript + esbuild セットアップ
npm run build          # TypeScript チェック

# テスト実行（VBA コンパイラ環境）
./node_modules/.bin/esbuild tests/spec/xxx.test.ts \
  --bundle --outfile=tests/spec/xxx.cjs --platform=node \
  && node tests/spec/xxx.cjs
```

### テストファイル構成

```
tests/
├── spec/                           # VBA コンパイラテスト
│   ├── math-module.test.ts        # Domain Logic テスト
│   ├── business-logic.test.ts     # 計算ロジック
│   └── edge-cases.test.ts         # エッジケース
└── ts/
    ├── test-runner.ts             # VBATest クラス
    └── sandbox.ts                 # ファイルシステム
```

---

## パターン1: 基本的な単体テスト

### VBA コード

```vba
' src/vba/math.vba
Function Sum(a As Long, b As Long) As Long
    Sum = a + b
End Function

Function Multiply(a As Long, b As Long) As Long
    Multiply = a * b
End Function
```

### TypeScript テスト（Jest パターン）

```typescript
// tests/spec/math.test.ts
import { VBATest } from '../ts/test-runner';

describe('Math Operations', () => {
  const vbaTest = new VBATest('src/vba/math.vba');

  describe('Sum', () => {
    it('should add two positive numbers', () => {
      const result = vbaTest.run('Sum', [10, 20]);
      expect(result).toBe(30);
    });

    it('should add negative numbers', () => {
      const result = vbaTest.run('Sum', [-10, -20]);
      expect(result).toBe(-30);
    });

    it('should handle zero', () => {
      const result = vbaTest.run('Sum', [0, 0]);
      expect(result).toBe(0);
    });
  });

  describe('Multiply', () => {
    it('should multiply two numbers', () => {
      const result = vbaTest.run('Multiply', [5, 4]);
      expect(result).toBe(20);
    });

    it('should handle multiplication by zero', () => {
      const result = vbaTest.run('Multiply', [100, 0]);
      expect(result).toBe(0);
    });
  });
});
```

**実行方法**:

```bash
./node_modules/.bin/esbuild tests/spec/math.test.ts \
  --bundle --outfile=tests/spec/math.cjs --platform=node \
  && node tests/spec/math.cjs
```

---

## パターン2: パラメータ化テスト

### VBA コード

```vba
' src/vba/discount.vba
Function ApplyDiscount(price As Double, discountPercent As Double) As Double
    If discountPercent < 0 Or discountPercent > 100 Then
        Err.Raise 5  ' Invalid argument
    End If
    ApplyDiscount = price * (1 - discountPercent / 100)
End Function
```

### TypeScript テスト

#### 方式1: 手動ループ（シンプル）

```typescript
describe('Discount Calculation', () => {
  const vbaTest = new VBATest('src/vba/discount.vba');

  const testCases = [
    { price: 1000, discount: 0, expected: 1000 },
    { price: 1000, discount: 10, expected: 900 },
    { price: 1000, discount: 50, expected: 500 },
    { price: 1000, discount: 100, expected: 0 },
    { price: 500, discount: 20, expected: 400 },
  ];

  testCases.forEach(({ price, discount, expected }) => {
    it(`${discount}% discount on $${price} = $${expected}`, () => {
      const result = vbaTest.run('ApplyDiscount', [price, discount]);
      expect(result).toBe(expected);
    });
  });

  it('should reject invalid discount', () => {
    expect(() => {
      vbaTest.run('ApplyDiscount', [1000, -10]);
    }).toThrow();
  });
});
```

#### 方式2: Jest `test.each`（Jest ネイティブ）

```typescript
describe('Discount Calculation with test.each', () => {
  const vbaTest = new VBATest('src/vba/discount.vba');

  // Jest 標準：test.each
  test.each([
    [1000, 0, 1000],
    [1000, 10, 900],
    [1000, 50, 500],
    [1000, 100, 0],
    [500, 20, 400],
  ])('discount %p%% on $%i = $%i', (price, discount, expected) => {
    const result = vbaTest.run('ApplyDiscount', [price, discount]);
    expect(result).toBe(expected);
  });
});
```

**メリット**: Jest のレポートが見やすい

---

## パターン3: 配列・複合データのテスト

### VBA コード

```vba
' src/vba/array-ops.vba
Function SumArray(arr As Variant) As Long
    Dim i As Integer
    Dim total As Long
    For i = LBound(arr) To UBound(arr)
        total = total + arr(i)
    Next i
    SumArray = total
End Function

Function FilterPositive(arr As Variant) As Variant
    Dim result() As Long
    Dim count As Integer
    Dim i As Integer
    
    ReDim result(0)
    count = 0
    
    For i = LBound(arr) To UBound(arr)
        If arr(i) > 0 Then
            ReDim Preserve result(count)
            result(count) = arr(i)
            count = count + 1
        End If
    Next i
    
    FilterPositive = result
End Function
```

### TypeScript テスト

```typescript
describe('Array Operations', () => {
  const vbaTest = new VBATest('src/vba/array-ops.vba');

  describe('SumArray', () => {
    it('should sum array of numbers', () => {
      const result = vbaTest.run('SumArray', [[10, 20, 30, 40]]);
      expect(result).toBe(100);
    });

    it('should handle negative numbers', () => {
      const result = vbaTest.run('SumArray', [[-10, 20, -30, 40]]);
      expect(result).toBe(20);
    });

    it('should return 0 for empty array', () => {
      const result = vbaTest.run('SumArray', [[]]);
      expect(result).toBe(0);
    });

    it('should handle single element', () => {
      const result = vbaTest.run('SumArray', [[42]]);
      expect(result).toBe(42);
    });
  });

  describe('FilterPositive', () => {
    it('should filter positive numbers', () => {
      const result = vbaTest.run('FilterPositive', [[10, -5, 20, -3, 30]]);
      expect(result).toEqual([10, 20, 30]);
    });

    it('should return empty array if no positive numbers', () => {
      const result = vbaTest.run('FilterPositive', [[-1, -2, -3]]);
      expect(result).toEqual([]);
    });

    it('should preserve order', () => {
      const result = vbaTest.run('FilterPositive', [[5, 1, 10, 2, 3]]);
      expect(result).toEqual([5, 1, 10, 2, 3]);
    });
  });
});
```

---

## パターン4: エラーハンドリングのテスト

### VBA コード

```vba
' src/vba/error-handling.vba
Function SafeDivide(a As Long, b As Long) As Double
    If b = 0 Then
        Err.Raise 11  ' Division by zero
    End If
    SafeDivide = a / b
End Function

Function SafeDivideWithDefault(a As Long, b As Long, defaultValue As Double) As Double
    On Error GoTo ErrorHandler
    SafeDivideWithDefault = a / b
    Exit Function
ErrorHandler:
    SafeDivideWithDefault = defaultValue
End Function
```

### TypeScript テスト

```typescript
describe('Error Handling', () => {
  const vbaTest = new VBATest('src/vba/error-handling.vba');

  describe('SafeDivide', () => {
    it('should divide normally', () => {
      const result = vbaTest.run('SafeDivide', [10, 2]);
      expect(result).toBe(5);
    });

    it('should throw on division by zero', () => {
      expect(() => {
        vbaTest.run('SafeDivide', [10, 0]);
      }).toThrow();
    });

    it('should handle negative divisor', () => {
      const result = vbaTest.run('SafeDivide', [10, -2]);
      expect(result).toBe(-5);
    });
  });

  describe('SafeDivideWithDefault', () => {
    it('should return default on error', () => {
      const result = vbaTest.run('SafeDivideWithDefault', [10, 0, -1]);
      expect(result).toBe(-1);
    });

    it('should divide normally without error', () => {
      const result = vbaTest.run('SafeDivideWithDefault', [10, 2, -1]);
      expect(result).toBe(5);
    });
  });
});
```

---

## パターン5: 複雑なビジネスロジック

### VBA コード

```vba
' src/vba/sales-logic.vba
Type SalesData
    Month As String
    Amount As Long
    Category As String
End Type

Function CalculateCommission(sales As Long, category As String) As Long
    Select Case category
        Case "A"
            CalculateCommission = CLng(sales * 0.1)
        Case "B"
            CalculateCommission = CLng(sales * 0.05)
        Case Else
            CalculateCommission = 0
    End Select
End Function

Function ProcessMultipleSales(amounts() As Long) As Long
    Dim i As Integer
    Dim total As Long
    For i = LBound(amounts) To UBound(amounts)
        If amounts(i) > 100 Then
            total = total + amounts(i)
        End If
    Next i
    ProcessMultipleSales = total
End Function
```

### TypeScript テスト

```typescript
describe('Sales Business Logic', () => {
  const vbaTest = new VBATest('src/vba/sales-logic.vba');

  describe('CalculateCommission', () => {
    // テストケースマトリックス
    const commissionTests = [
      { sales: 1000, category: 'A', expected: 100 },
      { sales: 1000, category: 'B', expected: 50 },
      { sales: 1000, category: 'C', expected: 0 },
      { sales: 0, category: 'A', expected: 0 },
      { sales: 5000, category: 'A', expected: 500 },
    ];

    commissionTests.forEach(({ sales, category, expected }) => {
      it(`category ${category}: ${sales} sales = ${expected} commission`, () => {
        const result = vbaTest.run('CalculateCommission', [sales, category]);
        expect(result).toBe(expected);
      });
    });
  });

  describe('ProcessMultipleSales', () => {
    it('should sum amounts over 100', () => {
      const result = vbaTest.run('ProcessMultipleSales', [
        [50, 150, 200, 30, 100]  // 150 + 200 = 350
      ]);
      expect(result).toBe(350);
    });

    it('should return 0 if no amounts exceed 100', () => {
      const result = vbaTest.run('ProcessMultipleSales', [
        [50, 75, 99]
      ]);
      expect(result).toBe(0);
    });

    it('should handle single large amount', () => {
      const result = vbaTest.run('ProcessMultipleSales', [
        [5000]
      ]);
      expect(result).toBe(5000);
    });
  });
});
```

---

## パターン6: ファイルシステム操作（VFS 使用）

### VBA コード

```vba
' src/vba/file-ops.vba
Function ReadAndProcessFile(filePath As String) As String
    ' VBA コンパイラ上では、仮想ファイルシステムから読み込み
    ' Domain Logic: ファイル内容を処理するだけ
    ReadAndProcessFile = UCase(filePath)  ' 例
End Function
```

### TypeScript テスト（VFS 活用）

```typescript
describe('File Operations with VFS', () => {
  const vbaTest = new VBATest('src/vba/file-ops.vba', {
    useVirtualFS: true
  });

  it('should read virtual file', () => {
    // VFS 上にテストデータを配置
    vbaTest.evaluator.fs.writeFileSync(
      '/sandbox/c/test-data.txt',
      'test content'
    );

    // VBA 関数で読み込み・処理
    const result = vbaTest.run('ReadFile', ['/sandbox/c/test-data.txt']);
    expect(result).toContain('test');
  });

  it('should handle missing file', () => {
    expect(() => {
      vbaTest.run('ReadFile', ['/sandbox/c/nonexistent.txt']);
    }).toThrow('File not found');
  });
});
```

---

## Jest の活用：便利な機能

### 1. `beforeEach` / `afterEach` でセットアップ・クリーンアップ

```typescript
describe('Database-like Operations', () => {
  let vbaTest: VBATest;

  beforeEach(() => {
    vbaTest = new VBATest('src/vba/db.vba');
    // 各テスト前にリセット
  });

  afterEach(() => {
    // 各テスト後にクリーンアップ
    vbaTest.evaluator.clearEnvironment();
  });

  it('should initialize correctly', () => {
    const result = vbaTest.run('Initialize', []);
    expect(result).toBe(true);
  });
});
```

### 2. `describe.skip` / `it.skip` で一時的に無効化

```typescript
describe('Feature Under Development', () => {
  // 開発中のテストは一時的に skip
  it.skip('should implement new feature', () => {
    // TODO: まだ実装されていない
  });

  it('should work for existing feature', () => {
    // こちらは実行される
  });
});
```

### 3. `test.each` で複数パターンを効率的にテスト

```typescript
test.each([
  [1, 2, 3],
  [2, 3, 5],
  [5, 5, 10],
  [-1, -2, -3],
])('Sum %i + %i = %i', (a, b, expected) => {
  const result = vbaTest.run('Sum', [a, b]);
  expect(result).toBe(expected);
});
```

### 4. スナップショットテスト（複雑な出力用）

```typescript
it('should generate expected report', () => {
  const result = vbaTest.run('GenerateReport', []);
  expect(result).toMatchSnapshot();
});
```

初回実行時にスナップショットを作成、以降の変更を検出できます。

---

## テスト実行の自動化

### npm scripts に追加

```json
{
  "scripts": {
    "test": "node tests/run-all-tests.js",
    "test:math": "esbuild tests/spec/math.test.ts --bundle --outfile=tests/spec/math.cjs --platform=node && node tests/spec/math.cjs",
    "test:watch": "nodemon --watch src/vba --exec 'npm run test'"
  }
}
```

### 一括実行スクリプト

```bash
#!/bin/bash
# tests/run-all-tests.js

const fs = require('fs');
const { execSync } = require('child_process');
const path = require('path');

const testDir = path.join(__dirname, 'spec');
const testFiles = fs.readdirSync(testDir)
  .filter(f => f.endsWith('.test.ts'));

let passed = 0;
let failed = 0;

testFiles.forEach(file => {
  console.log(`\n📝 Running ${file}...`);
  try {
    execSync(
      `./node_modules/.bin/esbuild tests/spec/${file} ` +
      `--bundle --outfile=tests/spec/${file.replace('.ts', '.cjs')} ` +
      `--platform=node && node tests/spec/${file.replace('.ts', '.cjs')}`,
      { stdio: 'inherit' }
    );
    passed++;
  } catch (error) {
    console.error(`❌ ${file} failed`);
    failed++;
  }
});

console.log(`\n\n📊 Results: ${passed} passed, ${failed} failed`);
process.exit(failed > 0 ? 1 : 0);
```

---

## ベストプラクティス

### 1. テストファイルは機能ごとに分割

```
tests/spec/
├── sales-calculation.test.ts
├── discount-logic.test.ts
├── error-handling.test.ts
└── edge-cases.test.ts
```

### 2. 説明的なテスト名を使用

```typescript
// ❌ 曖昧
it('works', () => { ... });

// ✅ 明確
it('should apply 10% discount to $100 and return $90', () => { ... });
```

### 3. AAA（Arrange-Act-Assert）パターン

```typescript
it('should calculate total correctly', () => {
  // Arrange: テスト条件の準備
  const input = [100, 200, 300];
  
  // Act: 実行
  const result = vbaTest.run('SumArray', [input]);
  
  // Assert: 検証
  expect(result).toBe(600);
});
```

### 4. Domain Logic テストに集中

```typescript
// ✅ テストすべき
it('CalculateTotalFromArray([100, 200]) = 300', () => {
  const result = vbaTest.run('CalculateTotalFromArray', [[100, 200]]);
  expect(result).toBe(300);
});

// ❌ テストすべきではない（Excel I/O は VBA IDE で）
// it('should save to Sheet1.C1', () => { ... });
```

---

## よくある質問

### Q: Mock が必要な場合は？

**A**: 以下の場合のみ：
- 時間に依存する処理（`Now()`）
- ランダム値（`Rnd()`）
- 外部 API 呼び出し

Domain Logic が適切に分離されていれば、通常は mock 不要です。

### Q: VBA オブジェクト（Sheets など）をテストしたい

**A**: VBA IDE で手動テストしてください。VBA コンパイラでのテストは推奨しません。

詳細は `docs/TESTING_STRATEGY.md` を参照。

### Q: テストの実行速度が遅い

**A**: 改善方法：
1. `test.skip` で不要なテストを一時的に無効化
2. ファイルを機能ごとに分割（並列化可能）
3. VBA コンパイラのバンドル時間を短縮（esbuild キャッシュ）

---

## まとめ

| 項目 | 推奨 |
|------|------|
| **テストフレームワーク** | Jest（既存） |
| **テスト対象** | Domain Logic のみ |
| **入力データ** | 配列、スカラー値（シンプル） |
| **Mock** | Domain Logic には不要 |
| **テスト実行** | npm scripts から |

**重要**:
- `docs/TESTING_STRATEGY.md` で VBA 設計原則を理解
- このドキュメントで JavaScript 実装を学ぶ
- 2つを合わせて、効率的で保守性の高いテストを実現

