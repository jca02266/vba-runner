# VBA テスト：JS テストフレームワーク活用ガイド

## はじめに

このドキュメントは、**既存の JavaScript テストフレームワーク（Jest など）を活用して、VBA マクロの単体テストを効率的に書く方法**を説明します。

`docs/TESTING_STRATEGY.md` で説明した Domain Logic テストを、実装レベルで具体化するガイドです。

---

## 推奨構成

```
VBA Runner + VBARunner クラス
           ↓
JavaScript テストフレームワーク（Jest）
           ↓
Domain Logic の単体テスト（配列等）
```

**メリット**:
- 既存の Jest エコシステムを活用（すぐに使える）
- VBA Runnerと TypeScript の統合
- 複雑な mock 不要（配列のみでテスト）

---

## 前提知識

**必ず先に読む**:
- `docs/TESTING_STRATEGY.md` - VBA テスト設計原則

本ドキュメントは、その原則を JavaScript 実装レベルで解説します。

---

## セットアップ

### 既存環境（推奨）

VBA Runner は既に以下を備えています：

```bash
# TypeScript + esbuild セットアップ
npm run build          # TypeScript チェック

# テスト実行（VBA Runner環境）
./node_modules/.bin/esbuild tests/spec/xxx.test.ts \
  --bundle --outfile=tests/spec/xxx.cjs --platform=node \
  && node tests/spec/xxx.cjs
```

### テストファイル構成

```
tests/
├── spec/                           # VBA Runnerテスト
│   ├── math-module.test.ts        # Domain Logic テスト
│   ├── business-logic.test.ts     # 計算ロジック
│   └── edge-cases.test.ts         # エッジケース
└── ts/
    ├── test-runner.ts             # VBARunner クラス
    └── sandbox.ts                 # ファイルシステム
```

---

## パターン1: 基本的な単体テスト [[→ T-01](REFACTORING_TESTING_CATALOG.md#t-01)]

### VBA コード

```vba
' src/vba/math.bas
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
import { VBARunner } from '../../test-libs/test-runner';

describe('Math Operations', () => {
  const vbaRunner = new VBARunner('src/vba/math.bas');

  describe('Sum', () => {
    it('should add two positive numbers', () => {
      const result = vbaRunner.run('Sum', [10, 20]);
      expect(result).toBe(30);
    });

    it('should add negative numbers', () => {
      const result = vbaRunner.run('Sum', [-10, -20]);
      expect(result).toBe(-30);
    });

    it('should handle zero', () => {
      const result = vbaRunner.run('Sum', [0, 0]);
      expect(result).toBe(0);
    });
  });

  describe('Multiply', () => {
    it('should multiply two numbers', () => {
      const result = vbaRunner.run('Multiply', [5, 4]);
      expect(result).toBe(20);
    });

    it('should handle multiplication by zero', () => {
      const result = vbaRunner.run('Multiply', [100, 0]);
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

## パターン2: パラメータ化テスト [[→ T-02](REFACTORING_TESTING_CATALOG.md#t-02)]

### VBA コード

```vba
' src/vba/discount.bas
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
  const vbaRunner = new VBARunner('src/vba/discount.bas');

  const testCases = [
    { price: 1000, discount: 0, expected: 1000 },
    { price: 1000, discount: 10, expected: 900 },
    { price: 1000, discount: 50, expected: 500 },
    { price: 1000, discount: 100, expected: 0 },
    { price: 500, discount: 20, expected: 400 },
  ];

  testCases.forEach(({ price, discount, expected }) => {
    it(`${discount}% discount on $${price} = $${expected}`, () => {
      const result = vbaRunner.run('ApplyDiscount', [price, discount]);
      expect(result).toBe(expected);
    });
  });

  it('should reject invalid discount', () => {
    expect(() => {
      vbaRunner.run('ApplyDiscount', [1000, -10]);
    }).toThrow();
  });
});
```

#### 方式2: Jest `test.each`（Jest ネイティブ）

```typescript
describe('Discount Calculation with test.each', () => {
  const vbaRunner = new VBARunner('src/vba/discount.bas');

  // Jest 標準：test.each
  test.each([
    [1000, 0, 1000],
    [1000, 10, 900],
    [1000, 50, 500],
    [1000, 100, 0],
    [500, 20, 400],
  ])('discount %p%% on $%i = $%i', (price, discount, expected) => {
    const result = vbaRunner.run('ApplyDiscount', [price, discount]);
    expect(result).toBe(expected);
  });
});
```

**メリット**: Jest のレポートが見やすい

---

## パターン3: 配列・複合データのテスト

### VBA コード

```vba
' src/vba/array-ops.bas
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
  const vbaRunner = new VBARunner('src/vba/array-ops.bas');

  describe('SumArray', () => {
    it('should sum array of numbers', () => {
      const result = vbaRunner.run('SumArray', [[10, 20, 30, 40]]);
      expect(result).toBe(100);
    });

    it('should handle negative numbers', () => {
      const result = vbaRunner.run('SumArray', [[-10, 20, -30, 40]]);
      expect(result).toBe(20);
    });

    it('should return 0 for empty array', () => {
      const result = vbaRunner.run('SumArray', [[]]);
      expect(result).toBe(0);
    });

    it('should handle single element', () => {
      const result = vbaRunner.run('SumArray', [[42]]);
      expect(result).toBe(42);
    });
  });

  describe('FilterPositive', () => {
    it('should filter positive numbers', () => {
      const result = vbaRunner.run('FilterPositive', [[10, -5, 20, -3, 30]]);
      expect(result).toEqual([10, 20, 30]);
    });

    it('should return empty array if no positive numbers', () => {
      const result = vbaRunner.run('FilterPositive', [[-1, -2, -3]]);
      expect(result).toEqual([]);
    });

    it('should preserve order', () => {
      const result = vbaRunner.run('FilterPositive', [[5, 1, 10, 2, 3]]);
      expect(result).toEqual([5, 1, 10, 2, 3]);
    });
  });
});
```

---

## パターン4: エラーハンドリングのテスト [[→ T-05](REFACTORING_TESTING_CATALOG.md#t-05)]

### VBA コード

```vba
' src/vba/error-handling.bas
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
  const vbaRunner = new VBARunner('src/vba/error-handling.bas');

  describe('SafeDivide', () => {
    it('should divide normally', () => {
      const result = vbaRunner.run('SafeDivide', [10, 2]);
      expect(result).toBe(5);
    });

    it('should throw on division by zero', () => {
      expect(() => {
        vbaRunner.run('SafeDivide', [10, 0]);
      }).toThrow();
    });

    it('should handle negative divisor', () => {
      const result = vbaRunner.run('SafeDivide', [10, -2]);
      expect(result).toBe(-5);
    });
  });

  describe('SafeDivideWithDefault', () => {
    it('should return default on error', () => {
      const result = vbaRunner.run('SafeDivideWithDefault', [10, 0, -1]);
      expect(result).toBe(-1);
    });

    it('should divide normally without error', () => {
      const result = vbaRunner.run('SafeDivideWithDefault', [10, 2, -1]);
      expect(result).toBe(5);
    });
  });
});
```

---

## パターン5: 複雑なビジネスロジック

### VBA コード

```vba
' src/vba/sales-logic.bas
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
  const vbaRunner = new VBARunner('src/vba/sales-logic.bas');

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
        const result = vbaRunner.run('CalculateCommission', [sales, category]);
        expect(result).toBe(expected);
      });
    });
  });

  describe('ProcessMultipleSales', () => {
    it('should sum amounts over 100', () => {
      const result = vbaRunner.run('ProcessMultipleSales', [
        [50, 150, 200, 30, 100]  // 150 + 200 = 350
      ]);
      expect(result).toBe(350);
    });

    it('should return 0 if no amounts exceed 100', () => {
      const result = vbaRunner.run('ProcessMultipleSales', [
        [50, 75, 99]
      ]);
      expect(result).toBe(0);
    });

    it('should handle single large amount', () => {
      const result = vbaRunner.run('ProcessMultipleSales', [
        [5000]
      ]);
      expect(result).toBe(5000);
    });
  });
});
```

---

## パターン6: ファイルシステム操作（VFS 使用） [[→ T-06](REFACTORING_TESTING_CATALOG.md#t-06)]

### VBA コード

```vba
' src/vba/file-ops.bas
Function ReadAndProcessFile(filePath As String) As String
    ' VBA Runner上では、仮想ファイルシステムから読み込み
    ' Domain Logic: ファイル内容を処理するだけ
    ReadAndProcessFile = UCase(filePath)  ' 例
End Function
```

### TypeScript テスト（VFS 活用）

```typescript
describe('File Operations with VFS', () => {
  const vbaRunner = new VBARunner('src/vba/file-ops.bas', {
    useVirtualFS: true
  });

  it('should read virtual file', () => {
    // VFS 上にテストデータを配置
    vbaRunner.evaluator.fs.writeFileSync(
      '/sandbox/c/test-data.txt',
      'test content'
    );

    // VBA 関数で読み込み・処理
    const result = vbaRunner.run('ReadFile', ['/sandbox/c/test-data.txt']);
    expect(result).toContain('test');
  });

  it('should handle missing file', () => {
    expect(() => {
      vbaRunner.run('ReadFile', ['/sandbox/c/nonexistent.txt']);
    }).toThrow('File not found');
  });
});
```

---

## パターン7: 日時依存テスト（Time Mocking） [[→ T-07](REFACTORING_TESTING_CATALOG.md#t-07)]

`Now()` や `Date()` を使う VBA コードは、実行するたびに結果が変わるため通常はテストできません。
`mockDate()` で日時を固定すれば、決定論的なテストが書けます。

> **設計上の注意**: `Now()` / `Date()` を関数内部で直接呼び出すのではなく、日時を引数として受け取る設計にすると `mockDate()` が不要になり、より純粋なテストが書けます。呼び出し箇所が多く変更が難しい場合や、リファクタリングの過渡期には `mockDate()` を活用してください。

### VBA コード

```vba
' src/vba/fiscal-year.bas
Function IsCurrentFiscalYear(targetYear As Long) As Boolean
    ' 現在年度（4月始まり）を判定
    Dim nowDate As Date
    nowDate = Now()
    Dim fiscalYear As Long
    If Month(nowDate) >= 4 Then
        fiscalYear = Year(nowDate)
    Else
        fiscalYear = Year(nowDate) - 1
    End If
    IsCurrentFiscalYear = (fiscalYear = targetYear)
End Function

Function DaysUntilDeadline(deadline As Date) As Long
    DaysUntilDeadline = DateDiff("d", Now(), deadline)
End Function
```

### TypeScript テスト

```typescript
import { VBARunner } from '../../test-libs/test-runner';

describe('Fiscal Year Logic', () => {
  let vbaRunner: VBARunner;

  beforeEach(() => {
    vbaRunner = new VBARunner('src/vba/fiscal-year.bas');
  });

  afterEach(() => {
    // 他テストへの影響を防ぐためモックを解除
    vbaRunner.mockDate(null);
  });

  it('should detect current fiscal year (April–March)', () => {
    // 2024年4月1日（FY2024の開始）
    vbaRunner.mockDate('2024-04-01T00:00:00Z');
    expect(vbaRunner.run('IsCurrentFiscalYear', [2024])).toBe(true);
    expect(vbaRunner.run('IsCurrentFiscalYear', [2023])).toBe(false);
  });

  it('should treat March as previous fiscal year', () => {
    // 2024年3月31日はまだFY2023
    vbaRunner.mockDate('2024-03-31T00:00:00Z');
    expect(vbaRunner.run('IsCurrentFiscalYear', [2023])).toBe(true);
    expect(vbaRunner.run('IsCurrentFiscalYear', [2024])).toBe(false);
  });

  it('should calculate days until deadline', () => {
    vbaRunner.mockDate('2024-01-01T00:00:00Z');
    // 2024-01-11 までの日数
    const days = vbaRunner.run('DaysUntilDeadline', [new Date('2024-01-11')]);
    expect(days).toBe(10);
  });
});
```

### `mockDate` の仕様

| 呼び出し | 効果 |
|---------|------|
| `vbaRunner.mockDate('2024-03-15T10:30:45Z')` | `Now()` / `Date()` / `Time()` / `Timer` が固定値を返す |
| `vbaRunner.mockDate(null)` | モックを解除し、実際のシステム時刻に戻す |

固定される関数一覧：

| VBA 関数 | 返す値 |
|---------|-------|
| `Now()` | モック日時（日付 + 時刻） |
| `Date()` | モック日時の日付部分のみ |
| `Time()` | モック日時の時刻部分のみ |
| `Timer` | モック日時の午前0時からの経過秒数 |
| `Year()` / `Month()` / `Day()` / `Hour()` / `Minute()` / `Second()` | `Now()` 経由で固定値を使用 |

---

## VBARunner API リファレンス

`test-libs/test-runner.ts` に定義された公開 API の一覧です。

### コンストラクタ

```typescript
new VBARunner(pathOrDir: string, config?: {
    sandboxRoot?: string,    // ファイル操作のサンドボックスルート
    env?: Record<string, string>,  // VBA の Environ() が返す環境変数
    useVirtualFS?: boolean   // true: メモリ内VFS / false: 実FS（デフォルト: false）
})
```

- `pathOrDir` にファイルパスを渡すと単一ファイルを読み込む
- ディレクトリパスを渡すと `.bas` / `.cls` / `.frm` を**名前順**に全ロード
- `.cls` ファイルは自動的にクラスモジュールとして解釈される（`Class ... End Class` ラッパー不要）

### メソッド

#### `run(procedureName, args, type?)` → `any`

指定プロシージャを呼び出し、戻り値を返す。実行時間と結果を `console.log` に出力する。

```typescript
vbaRunner.run('CalcTax', [50000])       // 関数呼び出し
vbaRunner.run('MyValue', [], 'get')     // Property Get
vbaRunner.run('MyValue', [100], 'let')  // Property Let
vbaRunner.run('MyObj', [obj], 'set')    // Property Set
```

#### `eval(exprString)` → `any`

VBA の式や文を文字列で評価する。式なら戻り値を返し、文なら `undefined` を返す。

```typescript
vbaRunner.eval('MyGlobal = 42')        // 変数の設定（文）
const v = vbaRunner.eval('2 + 3')      // 式の評価 → 5
```

#### `set(name, value)` → `void`

グローバル変数に TypeScript の値を直接注入する。VBA で参照できる。

```typescript
vbaRunner.set('ExchangeRate', 150.5)
vbaRunner.set('TargetSheet', mockWorksheet)
```

#### `mockDate(dateStr | null)` → `void`

`Now` / `Date` / `Time` / `Timer` が返す日時を固定する。`null` で実時刻に戻す。
ISO 8601 文字列（`'2024-03-15T10:30:00Z'`）推奨。

```typescript
vbaRunner.mockDate('2024-12-31T23:59:59Z')  // 日時を固定
vbaRunner.mockDate(null)                      // 解除
```

#### `spy(name, returnFn?)` → `SpyRecord` [[→ T-08](REFACTORING_TESTING_CATALOG.md#t-08)]

VBA 関数をスパイでラップし、呼び出し記録を返す。`returnFn` を指定すると戻り値を上書きできる。

```typescript
// 基本: 呼び出し記録だけ（MsgBox の元の動作を維持）
const spy = vbaRunner.spy('MsgBox');
vbaRunner.run('MyProc', []);

spy.callCount              // 呼び出し回数
spy.calls                  // 引数配列の配列: [['msg1'], ['msg2', 0, 'Title'], ...]
spy.lastCall               // 最後の呼び出しの引数配列
spy.calledWith('Error!')   // 指定引数で呼ばれたか（boolean）
spy.returnValues           // 各呼び出しの戻り値
spy.reset()                // 履歴をリセット

// 戻り値モック: ユーザーが「はい」を選んだ状態を再現（vbYes = 6）
const spy2 = vbaRunner.spy('MsgBox', () => 6);

// InputBox をモックして固定文字列を返す
vbaRunner.spy('InputBox', () => 'Alice');
```

`SpyRecord` は `test-libs/test-runner` から型としてもインポートできる。

```typescript
import { VBARunner, SpyRecord } from '../../test-libs/test-runner';
// ※ SpyRecord は evaluator.ts から再エクスポートされていないため、
//    型だけ必要な場合は import type { SpyRecord } from '../../src/engine/evaluator';
```

**使用例: MsgBox のエラーメッセージを検証する**

```typescript
it('should show error when input is invalid', () => {
  const spy = vbaRunner.spy('MsgBox');
  
  vbaRunner.eval('Validate("")');  // 空文字 → エラー表示
  
  expect(spy.callCount).toBe(1);
  expect(spy.calledWith('入力値が不正です')).toBe(true);
});
```

#### `registerExternalObject(progId, factory)` → `void` [[→ T-09](REFACTORING_TESTING_CATALOG.md#t-09)]

`CreateObject(progId)` が返すオブジェクトをテスト用スタブに差し替える。

```typescript
import { createRegExpMock } from '../../test-libs/regexp-mock';
vbaRunner.registerExternalObject('VBScript.RegExp', createRegExpMock);
```

#### `getTypeDefinitions()` → `Record<string, Record<string, string>>`

VBA ソースに含まれる `Type` 宣言を TypeScript の型情報として返す。VBA の型は TypeScript の型文字列にマッピングされる。

```typescript
const vbaRunner = new VBARunner('src/vba/inventory.bas');
const types = vbaRunner.getTypeDefinitions();
// => { InventoryParams: { CurrentStock: 'number', SoldUnits: 'number', ... } }

// interface 文字列として出力する場合:
for (const [name, fields] of Object.entries(types)) {
    const body = Object.entries(fields).map(([f, t]) => `  ${f}: ${t};`).join('\n');
    console.log(`interface ${name} {\n${body}\n}`);
}
```

型マッピング: `Integer` / `Long` / `Single` / `Double` / `Currency` / `Decimal` / `Byte` / `Date` → `number`、`String` → `string`、`Boolean` → `boolean`、`Object` → `object`、それ以外 → `any`。

#### `evaluator` プロパティ

`Evaluator` インスタンスへの直接アクセス。高度な操作（ファイルシステム書き込み、環境変数設定など）に使う。

```typescript
// VFS にテスト用ファイルを配置
(vbaRunner.evaluator as any).fs.writeFileSync('/sandbox/c/data.txt', 'hello');
```

---

## Jest の活用：便利な機能

### 1. `beforeEach` / `afterEach` でセットアップ・クリーンアップ

```typescript
describe('Database-like Operations', () => {
  let vbaRunner: VBARunner;

  beforeEach(() => {
    vbaRunner = new VBARunner('src/vba/db.bas');
    // 各テスト前にリセット
  });

  afterEach(() => {
    // 各テスト後にクリーンアップ
    vbaRunner.evaluator.clearEnvironment();
  });

  it('should initialize correctly', () => {
    const result = vbaRunner.run('Initialize', []);
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
  const result = vbaRunner.run('Sum', [a, b]);
  expect(result).toBe(expected);
});
```

### 4. スナップショットテスト（複雑な出力用）

```typescript
it('should generate expected report', () => {
  const result = vbaRunner.run('GenerateReport', []);
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
  const result = vbaRunner.run('SumArray', [input]);
  
  // Assert: 検証
  expect(result).toBe(600);
});
```

### 4. Domain Logic テストに集中

```typescript
// ✅ テストすべき
it('CalculateTotalFromArray([100, 200]) = 300', () => {
  const result = vbaRunner.run('CalculateTotalFromArray', [[100, 200]]);
  expect(result).toBe(300);
});

// ❌ テストすべきではない（Excel I/O は VBA IDE で）
// it('should save to Sheet1.C1', () => { ... });
```

---

## よくある質問

### Q: Mock が必要な場合は？

**A**: 以下の場合のみ：

| ケース | 対応方法 |
|-------|---------|
| 時間に依存する処理（`Now()`） | `vbaRunner.mockDate('2024-01-01')` — **パターン7**参照 |
| 外部オブジェクト（`CreateObject`） | `vbaRunner.registerExternalObject(progId, factory)` |
| ランダム値（`Rnd()`） | テスト前に `Randomize 固定シード` を VBA 側で呼ぶ |

Domain Logic が適切に分離されていれば、通常は mock 不要です。

### Q: テスト失敗時にどの行でエラーが起きたか知りたい

**A**: VBA 実行時エラーは `err.basLine` プロパティと `(line N)` 付きメッセージを持ちます。

```typescript
try {
    vbaRunner.run('MyProc', []);
} catch (e: any) {
    console.error('エラー行:', e.basLine);   // 例: 12
    console.error('メッセージ:', e.message); // 例: "Type mismatch (line 12)"
    console.error('Err.Number:', e.number);  // 例: 13
}
```

エラー番号の一覧は [VBA Error 番号一覧（TODO.md）](../TODO.md#vba-エラー番号別の改善項目errnumber-対応) を参照してください。

### Q: VBA オブジェクト（Sheets など）をテストしたい

**A**: VBA IDE で手動テストしてください。VBA Runnerでのテストは推奨しません。

詳細は `docs/TESTING_STRATEGY.md` を参照。

### Q: テストの実行速度が遅い

**A**: 改善方法：
1. `test.skip` で不要なテストを一時的に無効化
2. ファイルを機能ごとに分割（並列化可能）
3. VBA Runnerのバンドル時間を短縮（esbuild キャッシュ）

---

## まとめ

| 項目 | 推奨 |
|------|------|
| **テストフレームワーク** | Jest（既存） |
| **テスト対象** | Domain Logic のみ |
| **入力データ** | 配列、スカラー値（シンプル） |
| **日時モック** | `vbaRunner.mockDate('2024-01-01T00:00:00Z')` |
| **副作用スパイ** | `vbaRunner.spy('MsgBox')` → `SpyRecord` |
| **外部オブジェクト** | `vbaRunner.registerExternalObject(progId, factory)` |
| **エラーデバッグ** | `e.basLine` / `e.number` / `e.message` |
| **テスト実行** | npm scripts から |

**重要**:
- `docs/TESTING_STRATEGY.md` で VBA 設計原則を理解
- このドキュメントで JavaScript 実装を学ぶ
- 2つを合わせて、効率的で保守性の高いテストを実現

