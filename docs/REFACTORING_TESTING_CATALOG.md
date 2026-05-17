# リファクタリング・テスト手法カタログ

各ドキュメントで繰り返し登場する手法を名前付きでまとめた参照用カタログです。詳細・実装例は各リンク先を参照してください。

---

## リファクタリング手法

<a id="r-01"></a>
### R-01: 純粋関数の切り出し

巨大な `Sub` から副作用のない計算・判定ロジックを独立した `Function` として切り出す手法。関数名がビジネスルールの名前になり、引数と戻り値だけで動作が完結するため単独でテスト可能になる。

**具体例: Excel オブジェクト依存の分離**
VBA における典型的な適用場面。`Range`・`Cells`・`Sheets` などの Excel オブジェクトへのアクセスを `Sub` に限定し、ビジネスロジックを Excel 非依存の `Function` として切り出す。これにより `Function` 単体を Excel なしでテストできる。対象箇所の検出には `vba-analyzer` の `excelMockTargets`・`excelAccessCount` が使える（[VA-01](#va-01)）。

- [TESTING_STRATEGY.md — 原則1](TESTING_STRATEGY.md#原則1-domain-logic-と-excel-io-の徹底的分離)
- [TESTING_STRATEGY.md — §5](TESTING_STRATEGY.md#5-excel-オブジェクト依存は-sub-に限定)
- [REFACTORING_GUIDE.md — 原則1・パターン1〜3](REFACTORING_GUIDE.md#原則-1-職責の分離separation-of-concerns)
- [REFACTORING_EXAMPLE.md](REFACTORING_EXAMPLE.md)
- [MOCK_GUIDE.md — Step 0](MOCK_GUIDE.md#step-0-モックが必要かどうかを確認する)

---

<a id="r-02"></a>
### R-02: マジックナンバーの定数化

コード中に散らばる即値（行番号・列番号・シート名など）を、意味のある名前の定数（`Const` や変数）として抽出する手法。設定シート管理（[R-09](#r-09)）や定数グループ化（[R-06](#r-06)）への足がかりとなる最初のステップ。

VBA では `Range()`・`Cells()`・`Sheets()` の引数が典型的な候補となる：

```vba
' Before
Cells(3, 5).Value = Sheets("Sheet2").Cells(row, 2).Value

' After
Const COL_STATUS  As Long = 5   ' ステータス列
Const COL_MASTER  As Long = 2   ' マスタ参照列
Const SHEET_MASTER As String = "Sheet2"
Cells(3, COL_STATUS).Value = Sheets(SHEET_MASTER).Cells(row, COL_MASTER).Value
```

> **注意: 絶対位置か相対位置かを確認する**: `obj.Cells(row, col)` の `row`・`col` の意味は `obj` によって変わる。`obj` が省略または `Worksheet` の場合はシート全体の絶対行・列番号、`obj` が `Range` の場合はその Range の左上を起点とした相対行・列番号になる。相対位置の場合は `OFFSET_ROW_HEADER`・`OFFSET_COL_STATUS` のようなプレフィックスで命名規則を統一しておくと、絶対位置の定数（`ROW_HEADER`・`COL_STATUS`）と区別しやすくなる。

> **注意: 同じ値でも別の名前で**: 同じ数値（例: `3`）が複数箇所に現れても、別シートの別の列を指している場合がある。最初は**値が同じでも別の名前で抽出**しておくことが無難。リファクタリングを進める中で同じ意味と判明した時点で統合できる。

---

<a id="r-03"></a>
### R-03: UDT パラメータ集約

ドメインとして一体のパラメータ群を `Type` 宣言にまとめて関数に渡す手法。関数シグネチャが簡潔になり、関連しない変数が紛れ込みにくくなる。関係のないパラメータを無理にまとめることは避ける。

- [TESTING_STRATEGY.md — §3](TESTING_STRATEGY.md#3-関連パラメータは-type-にまとめる)
- [REFACTORING_GUIDE.md — パターン3](REFACTORING_GUIDE.md#パターン-3-複雑なビジネスロジックudt-を使用)
- [REFACTORING_EXAMPLE.md — 改善1](REFACTORING_EXAMPLE.md#改善-1-関心の分離と構造化udt-導入)

---

<a id="r-04"></a>
### R-04: クラスモデル化

状態保持・バリデーション・複数の操作が必要になった場合に `Class` を導入する手法。VBA のクラスはコンストラクタへの引数渡し・継承・オーバーロードがなく制限が多いため、`Type` + モジュール関数で足りる間はクラス化を急がない。

- [TESTING_STRATEGY.md — §4](TESTING_STRATEGY.md#4-振る舞いを持たせたい場合はクラスを検討する)
- [REFACTORING_GUIDE.md — パターン4](REFACTORING_GUIDE.md#パターン-4-状態を持つ処理オブジェクト指向版)

---

<a id="r-05"></a>
### R-05: データ Double Buffering

Excel シートのデータをメモリ上の配列に一括読み込みし、配列だけで処理して最後に一括書き戻す手法。シートへのアクセス回数を最小化してパフォーマンスを向上し、ロジック部分を Excel 依存から切り離してテスト可能にする。

- [REFACTORING_EXAMPLE.md — フェーズ2 データ読み込み](REFACTORING_EXAMPLE.md#フェーズ-2-データ読み込み行-101-122)

---

<a id="r-06"></a>
### R-06: 定数グループ化

バラバラに定義された定数を意味のまとまりで `Type` にまとめ、初期化関数で一括セットアップする手法。変更箇所が一ヶ所に集中し、関連する定数の一貫性が保ちやすくなる。

- [REFACTORING_EXAMPLE.md — 改善1](REFACTORING_EXAMPLE.md#改善-1-関心の分離と構造化udt-導入)

---

<a id="r-07"></a>
### R-07: ホットスポット優先戦略

すべてのコードを一度にリファクタリングするのではなく、変更頻度が高い・複雑度が高い・テストが切実に必要な箇所から着手する戦略。投資対効果を最大化する。

- [REFACTORING_GUIDE.md — フェーズ1](REFACTORING_GUIDE.md#フェーズ-1-ホットスポットを特定)

---

<a id="r-08"></a>
### R-08: 段階的リファクタリング

1つの `Function` を抽出 → テスト追加 → 元の `Sub` から呼び出しに差し替え、という小さなサイクルを繰り返す手法。一度に大きく変えずリグレッションリスクを低く保つ。

- [REFACTORING_GUIDE.md — フェーズ2〜3](REFACTORING_GUIDE.md#フェーズ-2-小さく始める)

---

<a id="r-09"></a>
### R-09: セル位置の設定シート管理

列番号・行番号の即値をコードに埋め込まず、Excel の設定シート（テーブル）で管理して `GetSetting()` などの関数で動的に参照する手法。シートレイアウト変更時のコード修正を不要にする。

- [REFACTORING_EXAMPLE.md — セル位置の即値を避ける](REFACTORING_EXAMPLE.md#さらなるリファクタリングセル位置の即値を避ける)

---

<a id="r-10"></a>
### R-10: クロージャによるループ抽象化（Strategy パターン）

キーブレイクループのような「構造は共通・中身だけ違う」処理を、VBA では **インターフェース + クラス** で抽象化する手法。JavaScript のクロージャに相当し、Java の匿名クラスと同じアプローチ。ループ構造を汎用プロシージャに切り出し、処理の中身（各レコードへの処理・グループ終了時の処理）を `Implements` で差し替えることで再利用性を高める。

**典型例**: 部門別売上小計、キーブレイクによるグループ集計

- [REFACTORING_CLOSURE.md — 詳細と全コード例](REFACTORING_CLOSURE.md)

---

## テスト手法

<a id="t-01"></a>
### T-01: 純粋関数テスト

副作用なし・Excel 依存なしの `Function` を直接呼び出して入出力を検証するテストパターン。最もシンプルで安定しており、テスト設計の出発点となる。

- [TESTING_STRATEGY.md — パターン1](TESTING_STRATEGY.md#パターン1-純粋関数テスト最適)
- [TEST_FRAMEWORK_GUIDE.md — パターン1](TEST_FRAMEWORK_GUIDE.md#パターン1-基本的な単体テスト)

---

<a id="t-02"></a>
### T-02: パラメータ化テスト（forEach）

同じ検証ロジックに対して複数の入力・期待値ペアをまとめて `forEach` で実行するパターン。テストケースの追加が容易で、網羅性を高めやすい。

```typescript
const cases: [InputType, number, string][] = [
    [{ ... }, 120, '通常'],
    [{ ... },  10, 'MinStock に丸め'],
];
cases.forEach(([params, expected, label]) => {
    assert.strictEqual(calcFn(params), expected, label);
});
```

- [TESTING_STRATEGY.md — パターン2](TESTING_STRATEGY.md#パターン2-パラメータ化テスト複数ケース)
- [TEST_FRAMEWORK_GUIDE.md — パターン2](TEST_FRAMEWORK_GUIDE.md#パターン2-パラメータ化テスト)

---

<a id="t-03"></a>
### T-03: Partial + デフォルト値テストヘルパー

`Partial<T>` とデフォルト値のスプレッドを組み合わせたヘルパー関数を用意し、テストごとに変化するフィールドだけを書く手法。テストケースが多い場合にコードを大幅に短縮できる（TypeScript 上級者向け）。

```typescript
type Params = Parameters<typeof calcFn>[0];
const base: Params = { A: 0, B: 0, C: 10, D: 200 };
const calcWith = (overrides: Partial<Params>) => calcFn({ ...base, ...overrides });

assert.strictEqual(calcWith({ A: 100, B: 30 }), 120);
```

- [TESTING_STRATEGY.md — §3](TESTING_STRATEGY.md#3-関連パラメータは-type-にまとめる)

---

<a id="t-04"></a>
### T-04: 状態変更テスト

モジュール変数やクラスの状態を初期化してから複数回呼び出し、状態遷移を順番に検証するパターン。`beforeEach` で状態をリセットし、テスト間の干渉を防ぐ。

- [TESTING_STRATEGY.md — パターン3](TESTING_STRATEGY.md#パターン3-状態変更を伴うロジックcontroller)
- [TEST_FRAMEWORK_GUIDE.md — beforeEach/afterEach](TEST_FRAMEWORK_GUIDE.md#1-beforeeach--aftereach-でセットアップクリーンアップ)

---

<a id="t-05"></a>
### T-05: エラーハンドリングテスト

意図的に異常系の入力を与え、期待通りのエラーが発生することを検証するパターン。正常系と同じ粒度でエラー系のケースを網羅する。

- [TESTING_STRATEGY.md — パターン4](TESTING_STRATEGY.md#パターン4-エラーハンドリング)
- [TEST_FRAMEWORK_GUIDE.md — パターン4](TEST_FRAMEWORK_GUIDE.md#パターン4-エラーハンドリングのテスト)

---

<a id="t-06"></a>
### T-06: 仮想ファイルシステムテスト（VFS）

`MemoryFileSystem` を使い、実際のファイルを作成せずにファイル I/O を含む VBA コードをテストする手法。サンドボックスルートを設定してパステストも可能。

- [TEST_FRAMEWORK_GUIDE.md — パターン6](TEST_FRAMEWORK_GUIDE.md#パターン6-ファイルシステム操作vfs-使用)

---

<a id="t-07"></a>
### T-07: 日時モック（Time Mocking）

`vbaRunner.mockDate()` で `Now` / `Date` / `Time` / `Timer` が返す値を固定し、日時依存のロジックを決定論的にテストする手法。

```typescript
vbaRunner.mockDate('2024-12-31T09:00:00');
// ...テスト実行...
vbaRunner.mockDate(null); // 解除
```

- [TEST_FRAMEWORK_GUIDE.md — パターン7](TEST_FRAMEWORK_GUIDE.md#パターン7-日時依存テストtime-mocking)

---

<a id="t-08"></a>
### T-08: スパイ（Spy）

`vbaRunner.spy()` で VBA 関数をラップし、呼び出し回数・引数・戻り値を記録して検証する手法。`MsgBox` などの副作用のある関数の検証やモック化に使う。

```typescript
const spy = vbaRunner.spy('MsgBox', () => 6); // vbYes を返すモック
vbaRunner.run('MyProc', []);
assert.strictEqual(spy.callCount, 1);
assert.ok(spy.calledWith('確認しますか？'));
```

- [TEST_FRAMEWORK_GUIDE.md — spy()](TEST_FRAMEWORK_GUIDE.md#spyname-returnfn--spyrecord)

---

<a id="t-09"></a>
### T-09: 外部オブジェクトスタブ

`vbaRunner.registerExternalObject()` で `CreateObject(progId)` が返すオブジェクトをテスト用スタブに差し替える手法。`VBScript.RegExp` などのランタイム依存を排除できる。

- [TEST_FRAMEWORK_GUIDE.md — registerExternalObject()](TEST_FRAMEWORK_GUIDE.md#registerexternalobjectprogid-factory--void)
- [MOCK_GUIDE.md](MOCK_GUIDE.md)

---

<a id="t-10"></a>
### T-10: MockWorksheet 統合テスト

`MockApplication` / `MockWorksheet` を使い、Excel I/O を含む複数関数の相互作用を検証する統合テストパターン。単体テストでは確認できない関数間の連携を検証できる。

- [MOCK_GUIDE.md — Part 1](MOCK_GUIDE.md#1-クイックスタート)
- [INTEGRATION_TEST_EXAMPLE.md](INTEGRATION_TEST_EXAMPLE.md#なぜ統合テストが必要か)

---

<a id="t-11"></a>
### T-11: 手動モック実装（4パターン）

既存の `MockWorksheet` を使わず、テスト対象に合わせた最小限のモックオブジェクトを自作する手法。軽量オブジェクト・クラス・ビルダー・JSON ベースの 4 パターンがある。

- [MOCK_GUIDE.md — §7](MOCK_GUIDE.md#7-自作モックのパターン4-つ)

---

<a id="t-12"></a>
### T-12: VBA IDE 手動テスト

Excel I/O や UI 操作を伴う `Sub` は VBA IDE 内で実行し、シートの状態を目視確認する手法。自動テストの対象外となる処理の検証方法。

- [TESTING_STRATEGY.md — 原則3](TESTING_STRATEGY.md#原則3-excel-io-テストは-vba-ide-で行う)

---

## テスト戦略

<a id="s-01"></a>
### S-01: テストピラミッド戦略

Unit Tests（70〜80%）→ Integration Tests（15〜25%）→ E2E / VBA IDE 手動テスト（5〜10%）の比率でテストを構成する戦略。下層ほど高速・安定で、上層ほど現実に近いが遅く壊れやすい。

- [TESTING_STRATEGY.md](TESTING_STRATEGY.md)
- [MOCK_GUIDE.md — §12](MOCK_GUIDE.md#12-テスト戦略段階的なテスト化)

---

## vba-analyzer ツール

<a id="va-01"></a>
### VA-01: vba-analyzer による静的解析支援

`vba-analyzer` は VBA ソースを静的解析し、リファクタリング・テスト設計の手がかりとなる情報を JSON 形式で出力するツールです。

```bash
node test-libs/vba-analyzer.cjs src/vba/ --json
```

#### 主要出力フィールド

| フィールド | 内容 | 活用手法 |
|---|---|---|
| `excelAccessCount` / `excelAccessSamples` | Excel オブジェクトへの直接アクセスを含む関数の数とサンプル。パラメータなしで Excel に依存している関数はリファクタリング候補の筆頭 | [R-01](#r-01) |
| `prefixClusters` | 共通の接頭辞（例: `inv_Stock`, `inv_Min`）を持つ変数群のグルーピング候補。`Type` 化やクラス化の手がかりになる | [R-03](#r-03) / [R-04](#r-04) |
| `excelMockTargets` | `Sheets`, `Range`, `Application` などへのアクセスを含む Function の一覧。Sub への切り出し（R-01）の対象候補 | [R-01](#r-01) |
| `excelObjectsUsed` | VBA コード中で参照されている Excel オブジェクト名の一覧。モック注入が必要なオブジェクトを特定する | [T-10](#t-10) |

#### 使用例

```bash
# リファクタリング候補を確認
node test-libs/vba-analyzer.cjs src/vba/ --json | jq '.excelMockTargets'

# モック対象オブジェクトを確認
node test-libs/vba-analyzer.cjs src/vba/ --json | jq '.excelObjectsUsed'
```

#### `--gen-test-dir <dir>`: テスト用ソースの自動生成

```bash
node test-libs/vba-analyzer.cjs src/vba/ --gen-test-dir tests/generated/
```

VBA ソース中で参照されているが定義されていない `xl*` / `vb*` / `mso*` 形式の Excel/VBA 定数を検出し、指定ディレクトリに `const.ts` として書き出します。

```typescript
// tests/generated/const.ts（出力例）
export const xlUp = -4162;
export const xlDown = -4121;
export const vbYes = 6;
export const xlCalculationManual = 0; // TODO: 実際の値を設定

/** vbaRunner.setConstants(allConstants) で一括注入するためのオブジェクト */
export const allConstants: Record<string, number | string> = {
    xlUp, xlDown, vbYes, xlCalculationManual,
};
```

値が既知でない定数は `0` で出力され `// TODO` コメントが付きます。実際の値は [Microsoft Learn — Excel VBA 列挙型一覧](https://learn.microsoft.com/en-us/office/vba/api/overview/excel/) で確認してください。

**VBARunner への注入**: 生成した `allConstants` を `VBARunner.setConstants()` に渡すと、VBA コード内で `xlUp` などの定数が使えるようになります。

```typescript
import { VBARunner } from '../../test-libs/test-runner';
import { allConstants } from './generated/const';

const vbaRunner = new VBARunner('src/vba/MyMacro.bas');
vbaRunner.setConstants(allConstants);  // xl*/vb* 定数を一括注入

const result = vbaRunner.run('MyFunction', []);
```

#### テキスト出力: 即値引数の検出（定数化候補）

テキスト出力モードでは、`Range()`・`Cells()`・`Sheets()`・`Worksheets()` の引数に数値リテラル・文字列リテラルが使われている箇所をプロシージャ単位で報告します（[R-02](#r-02) / [R-09](#r-09) への足がかり）。同じ値・アクセス形式でまとめて何件・どの行番号に現れるかを表示します。

複数ファイルをまたいだ横断集計はなく、ファイル・プロシージャごとに独立して出力されます。同じ即値が複数ファイルに散らばっている場合は、各ファイルの出力を個別に確認してください。

```
[Function] LoadData  (12行, L10-L21)
    📌 即値引数（定数化候補）: 4種
      Cells(RowIndex=3): 2件 [L11, L15]
      Cells(ColumnIndex=5): 2件 [L11, L15]
      Sheets(Index="Config"): 1件 [L12]
      Worksheets(Index=2): 1件 [L14]
```

検出対象のアクセス形式:

| 形式 | 例 | 検出 |
|---|---|---|
| 直接呼び出し | `Cells(3, 5)` | ✅ |
| オブジェクト経由 | `ws.Cells(3, 5)` | ✅ |
| `.Item()` 形式 | `Cells.Item(3, 5)` / `ws.Cells.Item(3, 5)` | ✅ |
| Range 結果への直接インデックス | `Range("A1:B3")(3, 5)` | ✅ |
| Range 結果の `.Item()` | `Range("A1:B3").Item(3, 5)` | ✅ |
| `Worksheets()` | `Worksheets("Config")` / `Worksheets(1)` | ✅ |
| Range 変数経由 | `rng.Item(3, 5)` / `rng(3, 5)` | ❌ 型追跡が必要 |

---

<a id="va-02"></a>
### VA-02: ByRef パラメータへの代入検出（UDT 戻り値リファクタリング候補）

VBA では複数の値を返したいときに `ByRef` パラメータへの代入を使うパターンが多い。
しかし意図せず ByRef になっている（修飾子を省略すると VBA のデフォルトは ByRef）場合や、
出力パラメータが増えた結果として保守が難しくなっているケースがある。
`vba-analyzer` はプロシージャ内で ByRef パラメータに代入している箇所を検出し、
`Type` 構造体（UDT）を戻り値として返すリファクタリングの候補として報告する。

テキスト出力例:

```
[Sub] GetResult  (8行, L5-L12)
    ⚠️  ByRef パラメータへの代入（UDT 戻り値リファクタリング候補）:
      outValue As Long: 2件 [L7, L9]
      outFlag As Boolean: 1件 [L10]
```

**検出条件**:
- `ByRef` 明示 または修飾子なし（VBA デフォルトは ByRef）のパラメータ
- そのパラメータが代入文の左辺に現れる

**検出しないケース**:
- `ByVal` パラメータへの代入
- ByRef パラメータを読み取るだけで代入しない場合

**リファクタリング方針**:
出力パラメータが 2 つ以上ある場合、`Type` 構造体にまとめて `Function` の戻り値として返す設計が望ましい（[R-03](#r-03)）。
1 つだけの場合は `Function` の戻り値に昇格させるだけで済む。

---

<a id="va-03"></a>
### VA-03: テストひな形出力（`--gen-test-dir`）

`--gen-test-dir <dir>` オプションを指定すると、`const.ts`（定数ファイル）に加えて、
各 VBA ファイルに対応する Jest テストひな形 `<FileName>.test.ts` を生成する。

```bash
node test-libs/vba-analyzer.cjs src/vba/ --gen-test-dir tests/generated/
```

生成されるファイル:
- `tests/generated/const.ts` — xl*/vb* 定数の定義（[VA-01](#va-01)）
- `tests/generated/<FileName>.test.ts` — Function ごとのテストひな形

**純粋関数と非純粋関数の区別**:

| 条件 | 判定 | 出力 |
|---|---|---|
| Excel オブジェクト未使用 かつ I/O なし | 純粋 | `test(...)` — すぐ実装できる状態 |
| `Cells` / `Range` 等の Excel アクセスあり | 非純粋 | `test.skip(...)` + `// ⚠️` コメント |
| `MsgBox` / `InputBox` / `Debug.Print` あり | 非純粋 | `test.skip(...)` + `// ⚠️` コメント |

I/O 副作用の検出対象: `MsgBox`、`InputBox`、`Debug.Print`

**出力例**:

```typescript
// Auto-generated by vba-analyzer --gen-test-dir
import { VBARunner } from '../test-libs/test-runner';

const runner = new VBARunner('../src/vba/TaskScheduler_Core.bas');

describe('CalcNumRows', () => {
  test('基本動作', () => {
    const result = runner.run('CalcNumRows', [/* lastRow: number */, /* taskCfg: any */]) as number;
    expect(result).toBe(/* expected: number */);
  });
});

// ⚠️  GetLastTaskRow: impure — Excel access (Cells, Rows)
describe('GetLastTaskRow', () => {
  test.skip('基本動作', () => {
    const result = runner.run('GetLastTaskRow', [/* ws: any */, /* taskCfg: any */]) as number;
    expect(result).toBe(/* expected: number */);
  });
});
```

**注意点**:
- `Sub` は対象外（戻り値を持たないため）。`Function` のみひな形を生成する
- 引数の型コメントは VBA の型宣言から推定（`Long`/`Integer`/… → `number`、`String` → `string`、`Boolean` → `boolean`、その他 → `any`）
- 戻り値型も同様に推定して `as number` 等を付加する
- `test.skip` のブロックはモック設定を追加してから有効化する（[T-07](#t-07) 参照）
