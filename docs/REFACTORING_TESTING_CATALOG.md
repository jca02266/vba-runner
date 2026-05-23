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
### R-03: UDT パラメーター集約

ドメインとして一体のパラメーター群を `Type` 宣言にまとめて関数に渡す手法。関数シグネチャが簡潔になり、関連しない変数が紛れ込みにくくなる。関係のないパラメーターを無理にまとめることは避ける。

- [TESTING_STRATEGY.md — §3](TESTING_STRATEGY.md#3-関連パラメーターは-type-にまとめる)
- [REFACTORING_GUIDE.md — パターン3](REFACTORING_GUIDE.md#パターン-3-複雑なビジネスロジックudt-を使用)
- [REFACTORING_EXAMPLE.md — 改善1](REFACTORING_EXAMPLE.md#改善-1-関心の分離と構造化udt-導入)

---

<a id="r-04"></a>
### R-04: クラスモデル化

状態保持・バリデーション・複数の操作が必要になった場合に `Class` を導入する手法。VBA のクラスはコンストラクターへの引数渡し・継承・オーバーロードがなく制限が多いため、`Type` + モジュール関数で足りる間はクラス化を急がない。

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

<a id="r-11"></a>
### R-11: テーブル駆動パターン

大量の同じ構造を持つ `If-ElseIf` チェーンを、**データテーブル + 統一されたルックアップロジック**に置き換える手法。外側分岐 × 内側分岐の組み合わせがデシジョンテーブルとして表現できる場合に適用可能。列数が行ごとに異なる非対称テーブルも `threshold = 0` でスキップ可能。

**適用指標**:
- 外側分岐数 ≥ 3、内側分岐数 ≥ 2
- 各外側分岐が同じ構造を繰り返している
- 代入値が単純（計算式・副作用なし）

**自動検出**: `test-libs/table-driven-detector.ts` の `TableDrivenDetector` クラスで検出・スコアリング可能。

- [TABLE_DRIVEN_GUIDE.md — 完全ガイド](TABLE_DRIVEN_GUIDE.md)
- [REFACTORING_GUIDE.md — パターン5](REFACTORING_GUIDE.md#パターン-5-テーブル駆動パターン)

---

## テスト手法

<a id="t-01"></a>
### T-01: 純粋関数テスト

副作用なし・Excel 依存なしの `Function` を直接呼び出して入出力を検証するテストパターン。最もシンプルで安定しており、テスト設計の出発点となる。

- [TESTING_STRATEGY.md — パターン1](TESTING_STRATEGY.md#パターン1-純粋関数テスト最適)
- [TEST_FRAMEWORK_GUIDE.md — パターン1](TEST_FRAMEWORK_GUIDE.md#パターン1-基本的な単体テスト)

---

<a id="t-02"></a>
### T-02: パラメーター化テスト（forEach）

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

- [TESTING_STRATEGY.md — パターン2](TESTING_STRATEGY.md#パターン2-パラメーター化テスト複数ケース)
- [TEST_FRAMEWORK_GUIDE.md — パターン2](TEST_FRAMEWORK_GUIDE.md#パターン2-パラメーター化テスト)

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

- [TESTING_STRATEGY.md — §3](TESTING_STRATEGY.md#3-関連パラメーターは-type-にまとめる)

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

`vba-analyzer` は VBA ソースを静的解析し、リファクタリング・テスト設計の手がかりとなる情報を出力するツールです。単一ファイルとディレクトリの両方を対象にできます。

#### ビルドと基本実行

```bash
# ビルド（変更後に実行）
./node_modules/.bin/esbuild test-libs/vba-analyzer.ts --bundle --outfile=test-libs/vba-analyzer.cjs --platform=node

# 基本実行（テキスト形式）
node test-libs/vba-analyzer.cjs src/vba/
```

#### CLI オプション一覧

| オプション | 内容 | 主な用途 |
|---|---|---|
| _(なし)_ | 全項目をテキスト形式で出力 | ファイル単体を詳細確認するとき |
| `--json` | JSON 形式で出力 | `jq` で絞り込む・プログラムに渡すとき |
| `--summary-only` | エントリーポイント候補・Dead code・モック必要箇所・重複ブロックのみ表示 | ディレクトリ全体の俯瞰 |
| `--outline` | プロシージャ名と問題フラグだけのコンパクト要約 | AI に渡す前の絞り込み。出力が小さくトークン節約になる |
| `--commented-code` | コメントアウトされたコード候補のみ表示 | 死コードの確認。出力をそのまま AI に貼って判断させることを想定（[VA-04](#va-04)） |
| `--gen-test-dir <dir>` | テストひな形と定数ファイルを生成 | テスト整備の起点（[VA-03](#va-03)） |

#### 主要出力フィールド

| フィールド | 内容 | 活用手法 |
|---|---|---|
| `excelAccessCount` / `excelAccessSamples` | Excel オブジェクトへの直接アクセスを含む関数の数とサンプル。パラメーターなしで Excel に依存している関数はリファクタリング候補の筆頭 | [R-01](#r-01) |
| `prefixClusters` | 共通の接頭辞（例: `inv_Stock`, `inv_Min`）を持つ変数群のグルーピング候補。`Type` 化やクラス化の手がかりになる | [R-03](#r-03) / [R-04](#r-04) |
| `excelMockTargets` | `Sheets`, `Range`, `Application` などへのアクセスを含む Function の一覧。Sub への切り出し（R-01）の対象候補 | [R-01](#r-01) |
| `excelObjectsUsed` | VBA コード中で参照されている Excel オブジェクト名の一覧。モック注入が必要なオブジェクトを特定する | [T-10](#t-10) |
| `assignmentBlocks` | 5件以上連続する代入・Dim 宣言の塊。`shape` によって UDT 化・関数抽出などの手法を選ぶ | [VA-01b](#va-01b) |

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

#### テキスト出力: プロシージャごとの指標サマリー

テキスト出力では、各プロシージャの先頭に指標サマリー行と凝集度判定が表示されます。

```
[Function] LoadData  (L10-L91, refs=3)
    行数 82⚠️  / ネスト 4⚠️  / Excel 9⚠️
    凝集度: LOW  ❌  (Dim 18 / シート参照 3種 / アドレス 6種)  ← 変数多(18個), シート参照多(3種), アドレス多(6種)
    🧪 モック必要候補: Sheets, Range, ActiveSheet
    ⚠️  連続代入ブロック（関数抽出候補）:
      L15-L27: 13件 [shape:mostly-range-write] [root]
    ❌ ByRef パラメーターへの代入（UDT 戻り値リファクタリング候補）:
      outResult As Long: 2件 [L45, L60]
    ⚠️  繰り返し数値リテラル（定数化候補）: 12(×3), 5(×2)
    ⚠️  即値引数（定数化候補）: 4種
      Cells(RowIndex=3): 2件 [L11, L15]
      Sheets(Index="Config"): 1件 [L12]
```

**指標ごとの良し悪し判断基準**:

| 指標 | ✅ 良好 | ⚠️ 注意 | ❌ 要対応 |
|---|---|---|---|
| **行数** | 29行以下 | 30〜99行 | 100行以上 |
| **ネスト深さ** | 2以下 | 3〜4 | 5以上 |
| **Excel アクセス数** | 4以下 | 5〜9 | 10以上 |

> **Excel アクセス数**とは、プロシージャ内で `Sheets.xxx`・`Range.xxx`・`ActiveSheet.xxx` などの Excel オブジェクトへのメンバーアクセスが何回現れるかの回数です。テスト可能性の目安として使い、⚠️以上なら `🧪 モック必要候補` の行に具体的なオブジェクト名が出ます。

**凝集度判定（Dim 数・固定シート参照・固定アドレス引数の合計スコア）**:

| 指標 | 0点 | 1点 | 2点 | 3点 |
|---|---|---|---|---|
| Dim 数（変数・定数宣言の合計） | 0〜7 | 8〜14 | 15〜29 | 30以上 |
| 固定シート参照の種類数 (`Sheets("名前")`) | 0〜1 | 2 | 3〜4 | 5以上 |
| 固定アドレス引数の種類数 (`Range("A1")`, `Cells(r,c)`) | 0〜2 | 3〜4 | 5〜7 | 8以上 |

合計スコアで判定:

| スコア | 判定 | 意味 |
|---|---|---|
| 0 | **HIGH ✅** | 責務が絞られており、テスト・変更がしやすい |
| 1〜2 | **MED ⚠️** | やや複雑。肥大化する前に分割を検討 |
| 3以上 | **LOW ❌** | 多くの責務を抱えている。積極的なリファクタリング推奨 |

#### テキスト出力: 即値引数の検出（定数化候補）

テキスト出力モードでは、`Range()`・`Cells()`・`Sheets()`・`Worksheets()` の引数に数値リテラル・文字列リテラルが使われている箇所をプロシージャ単位で報告します（[R-02](#r-02) / [R-09](#r-09) への足がかり）。同じ値・アクセス形式でまとめて何件・どの行番号に現れるかを表示します。

複数ファイルをまたいだ横断集計はなく、ファイル・プロシージャごとに独立して出力されます。同じ即値が複数ファイルに散らばっている場合は、各ファイルの出力を個別に確認してください。

```
[Function] LoadData  (L10-L21, refs=0)
    行数 12✅ / ネスト 1✅ / Excel 0✅
    凝集度: HIGH ✅  (Dim 1 / シート参照 1種 / アドレス 2種)
    ⚠️  即値引数（定数化候補）: 4種
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

<a id="va-01b"></a>
### VA-01b: 連続代入ブロック検出（関数抽出・UDT化候補）

`AssignmentStatement`（代入）・`SetStatement`（Set 代入）・`VariableDeclaration`（Dim 宣言）・`ConstDeclaration`（Const 宣言）が **5件以上連続している箇所**を検出し、`shape`（形状）を付けて報告する。

テキスト出力例:

```
[Function] ConvertToJson  (L199-L455, refs=1)
    行数 257❌ / ネスト 6❌ / Excel 0✅
    凝集度: MED  ⚠️
    ⚠️  連続代入ブロック（関数抽出候補）:
      L200-L226: 26件 [shape:mostly-dim-decl] [root]
```

#### shape の種類と意味

| shape | 内容 | 典型的なリファクタリング手法 |
|---|---|---|
| `mostly-dim-decl` | `Dim` 宣言が大半を占める | 変数が多すぎるシグナル。関連する変数を **UDT（Type 構造体）にまとめる**か、変数が多い原因（処理が大きすぎる）を見てプロシージャを分割する（[R-03](#r-03)） |
| `mostly-range-write` | `Range`/`Cells` への書き込みが大半 | Excel への一括書き出し処理。ロジックから切り離して **専用の出力 Sub に抽出**する（[R-01](#r-01)）|
| `mostly-range-read` | `Range`/`Cells` からの読み込みが大半 | Excel からの一括読み込み処理。**専用の入力 Sub/Function に抽出**し、戻り値は UDT か配列で返す（[R-01](#r-01) / [R-03](#r-03)） |
| `mostly-assign` | 通常の変数への代入が大半 | 初期化処理や前処理の塊。**初期化専用の Sub/Function に抽出**して本体の行数を削減する |
| `mostly-set-obj` | `Set obj = ...` が大半 | オブジェクト生成・注入処理。**ファクトリ関数または初期化 Sub** として切り出す |
| `mostly-var-init` | 変数への初期値設定が大半 | `mostly-assign` と同様。定数化できるものは `Const` に昇格させる |
| `mixed` | 上記の混在 | 複数の責務が混在している可能性が高い。個別に内容を確認してから手法を選ぶ |

#### `mostly-dim-decl` の場合の判断基準

`Dim` 宣言が 10 件を超えていても、それ自体は必ずしも問題ではない。以下の観点で判断する:

| 状況 | 対処 |
|---|---|
| 変数名に共通の接頭辞がある（`inv_Stock`, `inv_Min` 等） | UDT にまとめる（[R-03](#r-03)）。`prefixClusters` も参照 |
| 変数が関数の複数フェーズ（読み込み・計算・書き込み）に分散して使われている | フェーズごとに関数を分割し、各フェーズに必要な変数だけを宣言する |
| ほとんどの変数が関数全体で使われており分割が難しい | まず ByRef 出力パラメーターを UDT にまとめ（[VA-02](#va-02)）、変数数を物理的に削減してから再評価する |

#### 検出しないケース

- 5件未満の連続（閾値未満はノイズとして除外）
- `If`/`For` などの制御構造をまたいだ非連続な代入（制御構造の内側は別ブロックとして独立に評価される）

---

<a id="va-02"></a>
### VA-02: ByRef パラメーターへの代入検出（UDT 戻り値リファクタリング候補）

VBA では複数の値を返したいときに `ByRef` パラメーターへの代入を使うパターンが多い。
しかし意図せず ByRef になっている（修飾子を省略すると VBA のデフォルトは ByRef）場合や、
出力パラメーターが増えた結果として保守が難しくなっているケースがある。
`vba-analyzer` はプロシージャ内で ByRef パラメーターに代入している箇所を検出し、
`Type` 構造体（UDT）を戻り値として返すリファクタリングの候補として報告する。

テキスト出力例:

```
[Sub] GetResult  (L5-L12, refs=1)
    行数 8✅ / ネスト 1✅ / Excel 0✅
    凝集度: HIGH ✅  (Dim 2 / シート参照 0種 / アドレス 0種)
    ❌ ByRef パラメーターへの代入（UDT 戻り値リファクタリング候補）:
      outValue As Long: 2件 [L7, L9]
      outFlag As Boolean: 1件 [L10]
```

**検出条件**:
- `ByRef` 明示または修飾子なし（VBA デフォルトは ByRef）のパラメーター
- そのパラメーターが代入文の左辺に現れる

**検出しないケース**:
- `ByVal` パラメーターへの代入
- ByRef パラメーターを読み取るだけで代入しない場合

**リファクタリング方針**:
出力パラメーターが 2 つ以上ある場合、`Type` 構造体にまとめて `Function` の戻り値として返す設計が望ましい（[R-03](#r-03)）。
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

---

<a id="va-04"></a>
### VA-04: コメントアウトされたコード検出（`--commented-code`）

`--commented-code` オプションを指定すると、連続するコメント行の中に VBA キーワードが多く含まれるブロックを抽出して表示する。

```bash
node test-libs/vba-analyzer.cjs src/vba/ --commented-code
```

#### 出力例

```
========================================
💬 コメントアウトされたコード候補
========================================

  [HIGH] TaskScheduler_Core.bas L45-L62  (18行, score=12.5)
    キーワード: Sub×2, Dim×4, If×3, End×3, For×1
    --- 内容 ---
    Sub OldCalculation()
        Dim i As Long
        Dim total As Double
        For i = 1 To 10
            If total > 0 Then
                ...
    End Sub

  [LOW] TaskScheduler_v1.bas L4-L44  (41行, score=3.8)
    キーワード: End×1
    --- 内容 ---
    【仕様説明】 - 自動スケジュールロジック v2.0
    ...
```

#### 信頼度の基準

| 信頼度 | スコア | 意味 |
|---|---|---|
| **HIGH** | 8 以上 | VBA キーワードが多数。コードである可能性が高い |
| **MEDIUM** | 4〜7 | 複数のキーワードが含まれる。人間またはAIによる確認推奨 |
| **LOW** | 2〜3 | キーワードは少ないが何らかの構造を含む。自然言語の仕様説明も混入しやすい |

スコア 2 未満のブロックは出力されない。

#### スコアリング方式

| 要素 | 重み |
|---|---|
| `Sub` / `Function` / `Class` / `Property` | 3点/件 |
| `Type` | 2点/件 |
| `Dim` / `Const` / `Set` / `ReDim` / `If` / `For` / `Do` / `While` / `With` / `Select` / `Case` / `Call` / `End` / `Exit` / `GoTo` | 1点/件 |
| 代入文を含む行（`=` があり比較演算子でない） | 0.5点/行 |
| メンバーアクセスを含む行（`obj.prop` 形式） | 0.3点/行 |

#### 用途

- **ファイル全体を AI に渡す前の絞り込み**: `--commented-code` の出力だけを AI に貼り付けて「これはコードか説明文か」を判断させる。ファイル全体を渡すよりトークン消費を大幅に削減できる
- **Dead code の確認**: 過去にコメントアウトされたままのコードは削除候補。HIGH/MEDIUM のブロックを中心に確認する
- **復元候補の把握**: リファクタリング作業でコメントアウトされたが削除されていない旧実装を把握できる

#### 検出しないケース

- 2行以下の短いコメントブロック（ノイズとして除外）
- スコア 2 未満のブロック（自然言語の説明文として扱う）
- インラインコメント（`x = 1 ' 旧コード: x = 0` のような行末コメント）
