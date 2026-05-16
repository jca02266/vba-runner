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

### 全テスト実行

```bash
./run_all_tests.sh       # 簡潔表示
./run_all_tests.sh -v    # 詳細表示
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

### Step 1: 対象コードを把握する

```bash
# アウトライン表示（AI向けコンテキスト圧縮）
node test-libs/vba-analyzer.cjs sample/src/vba_legacy/ --outline

# 問題箇所の列挙
node test-libs/vba-analyzer.cjs sample/src/vba_legacy/TaskScheduler_v1.vba
```

### Step 2: リファクタリング前のスナップショットを取る

```bash
# 既存の参照テストが動くことを確認
./run_all_tests.sh
```

### Step 3: ロジックを純粋関数として抽出する

Excel依存箇所（`vba-analyzer` の `excelMockTargets` / `excelAccessSamples` で特定）を
インターフェース境界として扱い、純粋なビジネスロジック関数を別モジュールに抽出する。

パターン（`TaskScheduler_v1.vba` → `TaskScheduler_Core.vba` が実例）:
- セル読み込みブロック → 引数として受け取る関数に分離
- マジックナンバー → `Const` 宣言にまとめる（`prefixClusters` が候補を列挙）
- 巨大プロシージャ → 連続代入ブロック（`assignmentBlocks`）を切り出し関数化

### Step 4: テストを書いて検証する

`sample/tests/ts/` にテストを追加し、抽出した関数を実際に実行して検証する。
（Excel なしでロジックを検証できるのがこのプロジェクトの価値）

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
