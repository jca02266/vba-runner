# FOR_AI.md — このプロジェクトをリファクタリング支援で使うAIが最初に読むドキュメント

このプロジェクトは **VBAインタープリタ + 静的解析ツール** です。
Excelなしで VBA コードを実行・テスト・解析できます。
リファクタリング支援にのみ関与するAIは、このドキュメントを読めば他を読まずに作業できます。

---

## 読む必要のないディレクトリ・ファイル（スキップ推奨）

| パス | 理由 |
|---|---|
| `spec/` | MS-VBAL仕様書（4MBテキスト）。インタープリタ開発者向け。リファクタリングには不要 |
| `src/lsp/` | VSCode LSP拡張機能の実装。エディタ機能。リファクタリング作業とは無関係 |
| `src/App.tsx` 他 React ファイル | Web UI。リファクタリングには不要 |
| `dist/` | ビルド成果物 |
| `node_modules/` | パッケージ |
| `TODO.md` | VBAインタープリタのコンパイラ実装TODO。リファクタリング支援とは無関係 |
| `docs/TYPE_SYSTEM_SPEC.md` | インタープリタ内部型システムの仕様書 |
| `tests/lsp/`, `tests/engine/`, `tests/spec/` | インタープリタ自体のテスト。触らなくてよい |

---

## プロジェクトのコアファイル（リファクタリング支援に必要なもの）

```
src/compiler/
  lexer.ts          VBAソース → トークン列
  parser.ts         トークン列 → AST（全ノードに loc: {start, end} 付き）
  evaluator.ts      AST → 実行（VBAインタープリタ本体）
  sandbox.ts        ファイルI/Oのサンドボックス制限

test-libs/
  vba-analyzer.ts   静的解析CLI（リファクタリング支援の主ツール）★
  test-runner.ts    VBATest クラス（テスト実行ヘルパー）

sample/src/vba/
  TaskScheduler_Core.vba   純粋ビジネスロジック（Excelなし）
  TaskScheduler.vba        メインルーチン（Excel依存）
  Lib*.vba / Lib*.bas      ユーティリティライブラリ群

sample/src/vba_legacy/
  TaskScheduler_v1.vba     リファクタリング前の原典（394行の巨大関数）

sample/tests/ts/
  *.test.ts                VBAコードのユニットテスト群
```

---

## 主ツール1：静的解析 `vba-analyzer`

### ビルドと実行

```bash
# ビルド（初回、またはvba-analyzer.tsを変更したとき）
./node_modules/.bin/esbuild test-libs/vba-analyzer.ts --bundle --outfile=test-libs/vba-analyzer.cjs --platform=node

# 実行
node test-libs/vba-analyzer.cjs <ファイルまたはディレクトリ>          # テキスト出力
node test-libs/vba-analyzer.cjs <ファイルまたはディレクトリ> --json   # JSON出力（プログラム連携）
node test-libs/vba-analyzer.cjs <ファイルまたはディレクトリ> --outline # AI向けコンパクト要約
node test-libs/vba-analyzer.cjs <ファイルまたはディレクトリ> --summary-only  # ワークスペース集計のみ
```

### 検出できること

| 機能 | 出力 |
|---|---|
| プロシージャごとの行数・ネスト深さ・Dim数 | `lineCount`, `maxNestDepth`, `localDeclCount` |
| 問題フラグ | `LARGE(100行+)` `DEEP_NEST(5段+)` `MANY_LOCALS(30個+)` `EXCEL_HEAVY(10件+)` |
| 連続代入ブロック（抽出候補）+ 形状分類 | `assignmentBlocks[].shape` = `const-decl` / `dim-decl` / `range-read` / `range-write` / `var-init` / `assign` / `mixed` |
| Excel I/O アクセス箇所（**行番号付き**） | `excelAccessSamples[].line` |
| Excel モック必要候補（オブジェクト別） | `excelObjectsUsed[]` |
| 繰り返し数値リテラル（マジックナンバー） | `repeatedNumericLiterals[]` |
| 接頭辞クラスタ（UDT/Enum 抽出候補） | `prefixClusters[]` 例: `COL_×6` → Enum化提案 |
| エントリーポイント候補（Public・参照0） | `entryPointCandidates[]` + ヒューリスティック分類 |
| Dead code 候補（Private・参照0） | `deadCodeCandidates[]` |
| コールグラフ | `callGraph[].{from, fromFile, to, toFile}` |
| ワークスペース横断の参照カウント | `procedure.referenceCount` |

**注意**: Excel ボタン・イベントハンドラ・`Application.Run` による呼び出しは静的解析の範囲外。
参照0のPublicプロシージャを「dead code」として削除しないこと。

---

## 主ツール2：VBAインタープリタ（実行）

### テストの書き方

```typescript
// sample/tests/ts/MyFeature.test.ts
import { VBATest } from '../../test-libs/test-runner';

const vbaTest = new VBATest('sample/src/vba/MyModule.vba');

// サブルーチン呼び出し（引数あり）
const result = vbaTest.run('FunctionName', [arg1, arg2]);

// 式を評価
const val = vbaTest.eval('SomeSub()');  // 引数なし呼び出しにも使う
```

### テストのビルドと実行

```bash
./node_modules/.bin/esbuild sample/tests/ts/MyFeature.test.ts --bundle --outfile=sample/tests/ts/MyFeature.test.cjs --platform=node && node sample/tests/ts/MyFeature.test.cjs
```

> **注意**: `esbuild` は PATH に入っていないため、常に `./node_modules/.bin/esbuild` で実行すること。`npx` は使わない。

### Excelオブジェクトのモック

Excelオブジェクト（`ActiveSheet`, `Range`, `Cells` 等）はインタープリタで直接動作しない。
テストでは `VBATest` の第2引数にモックを渡す:

```typescript
const vbaTest = new VBATest('sample/src/vba/TaskScheduler_Core.vba', {
    sheets: mockSheets,
    activeSheet: mockSheet,
});
```

詳細: `docs/MOCK_GUIDE.md`

---

## リファクタリング支援のワークフロー

> **⚠️ 禁止事項: VBAファイルを `cat` / `Read` で全文読むことは禁止。**
> 全文読み込みはトークンを大量消費し、このツールの存在意義を消す。
> 常に `vba-analyzer` の出力を起点とし、**アナライザが示した行番号の範囲だけを狙い読みする**こと。

---

### Step 1: 対象コードを把握する（アナライザのみ使う）

```bash
# アウトラインで全体把握（全モジュール・全プロシージャを50行程度に圧縮）
node test-libs/vba-analyzer.cjs <対象ディレクトリ> --outline

# 詳細な問題箇所一覧（行番号付き）
node test-libs/vba-analyzer.cjs <対象ファイル>
```

アナライザ出力には **行番号が付いている**。コードを読む必要があれば、その行番号を使って範囲指定で読む:

```bash
# ✅ 正しい読み方: アナライザが示した行番号の前後だけ読む
sed -n '79,85p' target.vba

# ❌ 禁止: ファイル全体を読む
cat target.vba        # 禁止
Read target.vba       # 禁止（ツールの場合も同様）
```

---

### Step 2: テストで安全網を張る（リファクタリング前に必須）

**`./run_all_tests.sh` はインタープリタ自体のテストであり、対象VBAコードのテストではない。**
リファクタリング対象のVBAコードには、自分でテストを書く必要がある。

#### 2-1. 既存テストがある場合: まず GREEN を確認する

```bash
./node_modules/.bin/esbuild sample/tests/ts/MyFeature.test.ts \
  --bundle --outfile=sample/tests/ts/MyFeature.test.cjs --platform=node \
  && node sample/tests/ts/MyFeature.test.cjs
```

既存テストが GREEN でなければ、リファクタリングを始めてはいけない。

#### 2-2. 既存テストがない場合: リファクタリング前の挙動をテストに記録する

リファクタリング前の関数の入出力を `VBATest` で記録し、GREEN にしておく。
これが「壊れていないこと」を判断する唯一の根拠になる。

```typescript
// sample/tests/ts/MyFeature.test.ts
import { VBATest, assert } from '../../../test-libs/test-runner';

const vbaTest = new VBATest('sample/src/vba/MyModule.vba');
assert.strictEqual(vbaTest.run('TargetFunction', [input1, input2]), expected, 'description');
```

---

### Step 3: 1件だけ選んで抽出する

アナライザ出力はリファクタリング候補の一覧であり、**全部やる指示ではない**。

ユーザーが「何を改善したいか」を確認してから、候補の中で最も優先度が高い1件を選ぶ。
ユーザーが指定しない場合は、以下の観点で1件選びユーザーに提案する:
- **影響範囲が小さく、テストで検証しやすい**もの
- `duplicateBlocks` の stmtCount × occurrences が大きい（最もインパクトがある重複）
- `DEEP_NEST` のうち、内側ループが独立して意味を持つもの

選んだら、その行番号を `sed` で読んでから抽出する。

抽出の典型パターン:
- `duplicateBlocks` → 重複する N 行を1つの関数に切り出し、両箇所から呼ぶ
- `assignmentBlocks` → 連続代入ブロックをデータ構造の初期化関数に切り出す
- `DEEP_NEST` → 深いネストの内側ループをサブルーチンに切り出す

1件の抽出が終わったら、次に進む前にユーザーに確認する。
「続けますか？」が基本姿勢。自動的に全件処理しない。

---

### Step 4: 抽出した関数のテストを書く（抽出直後に必須）

**統合テストのGREENだけでは不十分。** 理由:

- 統合テスト（呼び出し元のテスト）は「正常系の代表データ」しか与えていない
- 抽出した関数が担う**エッジケース**（空値・非数値・境界値など）は統合テストの外にある
- 抽出でロジックを一箇所に集めた意味は、そこを直接テストして初めて完結する

```typescript
// 抽出した関数のエッジケースを直接テストする
console.log("\n[Test Suite] GetNumericCellValue");

// 通常値
assert.strictEqual(vbaTest.run('GetNumericCellValue', [grid, 1, 1]), 0.5, "numeric value");
// 空セル → 0
assert.strictEqual(vbaTest.run('GetNumericCellValue', [grid, 1, 2]), 0,   "empty cell → 0");
// 文字列 → 0
assert.strictEqual(vbaTest.run('GetNumericCellValue', [grid, 1, 3]), 0,   "string cell → 0");
// null → 0
assert.strictEqual(vbaTest.run('GetNumericCellValue', [grid, 1, 4]), 0,   "null cell → 0");
```

テストを追加したら再度実行して、既存テストと新テストがすべて GREEN になることを確認する。

---

### Step 5: アナライザで今回の変更を確認する

```bash
node test-libs/vba-analyzer.cjs <対象ファイル>
```

**今回選んだ1件**に対応するフラグが変化したことを確認する。
残っている他のフラグは「次回以降の候補」であり、今すぐ消す必要はない。

アナライザはあくまで候補の提示ツールであり、フラグをゼロにすることが目的ではない。
「全フラグを消す」を目指すと過剰なリファクタリングになる。止め時の判断はユーザーに委ねる。

---

## インタープリタの既知の制約（リファクタリング作業中に詰まったら確認）

- `ActiveWorkbook`, `ActiveSheet`, `Cells`, `Range`, `Sheets` 等の Excel オブジェクトは自動でモック化されない。テストで明示的に渡す必要がある。
- `CreateObject("ADODB.Connection")` 等の COM オブジェクトはスタブ扱い（エラーにはならないが動作しない）
- ファイルI/O（`Open`, `Close` 等）は `sandbox/` ディレクトリ以下に制限される
- VBA の `Type` 宣言（UDT）は実装済み。`Class` モジュールも実装済み
- `On Error Resume Next` 実装済み。`On Error GoTo` も実装済み

詳細な対応状況: `TODO.md`（「VBA ランタイム挙動」セクション）

---

## サンプルコードの位置づけ（参照用）

| ファイル | 用途 |
|---|---|
| `sample/src/vba_legacy/TaskScheduler_v1.vba` | リファクタリング**前**の原典。巨大関数(394行)。解析のベンチマーク対象 |
| `sample/src/vba/TaskScheduler_Core.vba` | リファクタリング**後**の純粋ロジック。テスト対象 |
| `sample/src/vba/TaskScheduler.vba` | リファクタリング後のメインルーチン（Excel依存部分） |
| `sample/src/vba_legacy/TaskScheduler.md` | リファクタリングの設計メモ |
| `docs/REFACTORING_EXAMPLE.md` | TaskSchedulerのリファクタリング手順の記録 |
