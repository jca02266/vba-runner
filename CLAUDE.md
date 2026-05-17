# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

**リファクタリング支援として関与する場合は `FOR_AI.md` を先に読むこと。**
VBA実行エンジンの実装に関与する場合はこのファイルを続けて読む。

## コマンド

**Web UI 開発サーバー:**
```bash
npm run dev        # http://localhost:5173/
npm run build      # TypeScript チェック + Vite ビルド
npm run lint       # ESLint
```

**テスト実行**（esbuild でバンドル後、Node で実行）:
```bash
./node_modules/.bin/esbuild sample/tests/ts/TaskScheduler_Core.test.ts --bundle --outfile=sample/tests/ts/TaskScheduler_Core.test.cjs --platform=node && node sample/tests/ts/TaskScheduler_Core.test.cjs
```

> **注意**: `esbuild` など Node.js のローカルツールは PATH に入っていないことを前提に、常に `./node_modules/.bin/<コマンド>` で直接実行すること。`npx` は使わない。

## アーキテクチャ

TypeScript で実装された VBA 実行エンジン。Excel 不要で VBA コードの実行・リファクタリング・ユニットテストを可能にする。

**コアパイプライン**（`src/engine/`）:
- `lexer.ts` — VBA ソースをトークン列に変換（50 種類以上のトークン型）
- `parser.ts` — トークンから AST を構築（ForStatement, IfStatement, ProcedureDeclaration など）
- `evaluator.ts` — `Environment` クラスでスコープ管理しながら AST を評価

**`Environment` スコープチェーン**: 変数とプロシージャは `enclosing` 参照を持つ `Environment` インスタンスで管理。VBA の大文字小文字無視の仕様に合わせ、識別子は小文字に正規化。未定義変数は暗黙的に `0` に初期化（VBA の仕様通り）。

**組み込み VBA 関数**（`evaluator.ts` に実装）: `IsEmpty`, `IsNumeric`, `CDBl`, `CLng`, `Int`, `UCase`, `Trim`, `UBound`, `CreateObject("Scripting.Dictionary")`（JS の `Map` で実装、内部アクセス用に `__map__` プロパティを持つ）。

**Web UI**（`src/App.tsx`）: リアルタイムのトークナイズによるシンタックスハイライト付きで、ブラウザ上で実行パイプラインを動かす React コンポーネント。

## テストパターン

テストは `sample/tests/ts/` に配置。`VBARunner` クラス（`test-libs/test-runner.ts`）はコンストラクタ時に `.vba` ファイルを読み込み、2 つの実行メソッドを提供する:

- **`vbaTest.run(procedureName, args)`** — TypeScript の配列引数を渡して名前付き Sub/Function を呼び出す
- **`vbaTest.eval(expressionString)`** — VBA の式または文を文字列として評価する。式の場合は戻り値を返し、文の場合は `undefined` を返す

引数付きのプロシージャ呼び出しには `run`、式の直接評価や引数なしのプロシージャ呼び出しには `eval` を使う。

## ファイル入出力の実装規則

**ファイル操作（Open, Close, Print#, Write#, Input# など §5.4.5 全般）を実装する際は、必ず README.md の「ファイル入出力のSandbox方針」を参照すること。**

実装上の必須要件:
- すべてのファイルパスは `src/engine/sandbox.ts` の `SandboxPath` クラスを通じて解決する
- Windowsパス（`C:\...`）はドライブレターをサブディレクトリとして変換する（例: `C:\foo` → `{sandboxRoot}/c/foo`）
- Sandboxルート外へのパス traversal（`../` など）は実行時エラーにする
- `Environ` 関数はOSの実環境変数を参照せず、Sandbox内の定義（メモリまたは `{sandboxRoot}/.env`）のみを返す
- テストコードでは `test-libs/sandbox.ts` のユーティリティでパスを組み立てる

## サンプルコードの構成

- `sample/src/vba/TaskScheduler_Core.vba` — 純粋なビジネスロジック関数（ユニットテスト対象）
- `sample/src/vba/TaskScheduler.vba` — リファクタリング後のメインルーチン（Excel オブジェクト依存のため直接テスト対象外）
- `sample/src/vba_legacy/TaskScheduler_v1.vba` — 参照用のリファクタリング前のオリジナルソース
