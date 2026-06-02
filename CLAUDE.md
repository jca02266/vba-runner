# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

**リファクタリング支援として関与する場合は `FOR_AI.md` を先に読むこと。**
VBA実行エンジンの実装に関与する場合はこのファイルを続けて読む。

## リポジトリ構成（成果物の分離）

ルートはエンジン本体（`src/`）と共通設定のみを持ち、**配布成果物ごとに `build/` 配下へ `package.json` を分離**している。`src/` は移動しないため、テスト・LSP・test-libs の import パスは不変。

| ディレクトリ | 成果物 | 主なソース |
|---|---|---|
| ルート `package.json` | エンジン本体 + 共通（test/lint/build 統括、全 devDeps） | `src/engine/` |
| `build/runner/` | npm パッケージ `vba-runner`（ライブラリ + CLI） | `test-libs/test-runner.ts`, `test-libs/vba-*.ts` |
| `build/extension/` | VS Code 拡張機能 (.vsix) | `src/extension.ts`, `src/lsp/` |
| `build/playground/` | Web UI デモ | `src/App.tsx`, `index.html` |

各配布物は esbuild `--bundle` で自己完結（Node 組み込みと extension の `vscode` external のみ）。`vba-runner` の `dependencies` は空（react 等の混入なし）。

## コマンド

**テスト実行**（tsx でそのまま実行）:
```bash
npx tsx tests/spec/eval-call-scope.test.ts
npm test            # tests/spec/ 一括（run_spec_tests.sh）
npm run typecheck   # tsc -b（プロジェクト参照を辿る。tsc --noEmit はルートが references のみで素通りするので不可）
npm run lint        # ESLint
```

**CLI ツールのローカル実行**（esbuild ビルド不要）:
```bash
npx tsx test-libs/vba-analyzer.ts <path>
npx tsx test-libs/vba-formatter.ts <path>
```

**成果物のビルド**（配布時）:
```bash
npm run build              # 3 成果物すべて（runner → extension → playground）
npm run build:runner       # build/runner/dist/{lib.cjs,bin/*.cjs}
npm run build:extension    # build/extension/dist/extension.cjs
npm run build:playground   # build/playground/dist/（Vite）
npm run package:extension  # build/extension/ で vsce package → .vsix
```

**Web UI 開発サーバー:**
```bash
npm run dev --prefix build/playground   # http://localhost:5173/
```

## 開発時実行とパッケージビルドの方針

| 用途 | 方法 | 備考 |
|------|------|------|
| 開発・テスト時 | `npx tsx <file>.ts` | esbuild 不要、ソースを直接実行 |
| npm パッケージ配布 | `npm run build:runner` → `build/runner/dist/*.cjs` | esbuild で CJS バンドルを生成 |
| 拡張機能配布 | `npm run package:extension` → `.vsix` | esbuild + vsce |

### `__dirname` の扱い

`package.json` が `"type":"module"` のため Node.js は ESM モードで動く。ソースコードでは ESM ネイティブの **`import.meta.dirname`** を使う。

```ts
// ✅ 正しい（ESM / tsx 両対応）
const dir = import.meta.dirname;

// ❌ 使わない（CJS 専用、ESM では ReferenceError）
const dir = __dirname;
```

CJS バンドル時（`npm run build`）は esbuild の `--define:import.meta.dirname=__dirname` オプションで自動的に `__dirname` へ置換されるため、配布後の `.cjs` ファイルも正常に動作する。

## アーキテクチャ

TypeScript で実装された VBA 実行エンジン。Excel 不要で VBA コードの実行・リファクタリング・ユニットテストを可能にする。

**コアパイプライン**（`src/engine/`）:
- `lexer.ts` — VBA ソースをトークン列に変換（50 種類以上のトークン型）
- `parser.ts` — トークンから AST を構築（ForStatement, IfStatement, ProcedureDeclaration など）
- `evaluator.ts` — `Environment` クラスでスコープ管理しながら AST を評価

**`Environment` スコープチェーン**: 変数とプロシージャは `enclosing` 参照を持つ `Environment` インスタンスで管理。VBA の大文字小文字無視の仕様に合わせ、識別子は小文字に正規化。未定義変数は暗黙的に `0` に初期化（VBA の仕様通り）。

**組み込み VBA 関数**（`evaluator.ts` に実装）: `IsEmpty`, `IsNumeric`, `CDBl`, `CLng`, `Int`, `UCase`, `Trim`, `UBound`, `CreateObject("Scripting.Dictionary")`（JS の `Map` で実装、内部アクセス用に `__map__` プロパティを持つ）。

**Web UI**（`src/App.tsx`）: リアルタイムのトークナイズによるシンタックスハイライト付きで、ブラウザ上で実行パイプラインを動かす React コンポーネント。

## テストディレクトリの使い分け

| ディレクトリ | 用途 |
|-------------|------|
| `tests/spec/` | VBA エンジンの言語仕様テスト。新機能実装・バグ修正時はここにテストを追加する |
| `tests/test-libs-tests/` | `test-libs/` 配下のテストランナー自身の機能テスト |
| `tests/spec/vba/` | 実 VBA 環境で動作する `.bas` 形式のテスト（`run_vba_tests.sh` で実行）|
| `sample/tests/ts/` | サンプル VBA コード（`sample/src/vba/`）のユニットテスト |

## VBA 仕様バグ修正・機能実装時の手順

VBA 仕様に関わる修正（バグ修正・未実装機能の追加）を行う場合は、**`.claude/commands/implement-vba.md` に記述された手順**に従うこと。要点：

1. `tests/spec/` にテストを追加する（`evalVBASingle` / `evalVBAModules` を使う）
2. `TODO_SPEC.md` の該当項目を更新する（`❌` → `✅` または `⚠️`）
3. 仕様バグ修正時は `TODO_SPEC.md` の「仕様バグ修正」セクションにも記録する

テストの書き方の詳細は `.claude/commands/implement-vba.md` の「Step 4: テストを作成」を参照。

## テストパターン

**`VBARunner`（`test-libs/test-runner.ts`）はアプリケーション向けのクラスであり、`tests/spec/` の低レベルなエンジンテストでは使わない。** `tests/spec/` では `evalVBASingle` / `evalVBAModules`（`test-libs/test-runner.ts` からエクスポート）を使う。

| テストの種類 | 使うもの | 配置先 |
|-------------|---------|--------|
| エンジン仕様テスト | `evalVBASingle` / `evalVBAModules` | `tests/spec/` |
| サンプル VBA のテスト | `VBARunner` | `sample/tests/ts/` |

テストは `sample/tests/ts/` に配置。`VBARunner` クラス（`test-libs/test-runner.ts`）はコンストラクター時に `.bas` ファイルを読み込み、2 つの実行メソッドを提供する:

- **`vbaRunner.run(procedureName, args)`** — TypeScript の配列引数を渡して名前付き Sub/Function を呼び出す
- **`vbaRunner.eval(expressionString)`** — VBA の式または文を文字列として評価する。式の場合は戻り値を返し、文の場合は `undefined` を返す

引数付きのプロシージャ呼び出しには `run`、式の直接評価や引数なしのプロシージャ呼び出しには `eval` を使う。

## ファイル入出力の実装規則

**ファイル操作（Open, Close, Print#, Write#, Input# など §5.4.5 全般）を実装する際は、必ず REFERENCE.md の「ファイル入出力のSandbox方針」を参照すること。**

実装上の必須要件:
- すべてのファイルパスは `src/engine/sandbox.ts` の `SandboxPath` クラスを通じて解決する
- Windowsパス（`C:\...`）はドライブレターをサブディレクトリとして変換する（例: `C:\foo` → `{sandboxRoot}/c/foo`）
- Sandboxルート外へのパス traversal（`../` など）は実行時エラーにする
- `Environ` 関数はOSの実環境変数を参照せず、Sandbox内の定義（メモリまたは `{sandboxRoot}/.env`）のみを返す
- テストコードでは `test-libs/sandbox.ts` のユーティリティでパスを組み立てる

## サンプルコードの構成

- `sample/src/refactoring/TaskScheduler_Core.bas` — 純粋なビジネスロジック関数（ユニットテスト対象）
- `sample/src/refactoring/TaskScheduler.bas` — リファクタリング後のメインルーチン（Excel オブジェクト依存のため直接テスト対象外）
- `sample/src/refactoring/TaskScheduler_v1.bas` — 参照用のリファクタリング前のオリジナルソース
