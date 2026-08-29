---
type: Reference
title: VBA Language Server 実装済み機能リファレンス
description: VS Code拡張機能とインラインLanguage Serverの実装仕様。
tags: [vba, lsp, vscode, implementation]
status: stable
---

# VBA Language Server — 実装済み機能リファレンス

> 対象: VS Code 拡張・LSP を開発・改修する人
>
> 前提: [CONTRIBUTING.md](../../CONTRIBUTING.md)、利用者向け機能一覧は [build/extension/README.ja.md](../../build/extension/README.ja.md)
>
> 次に読む: [REFERENCE.md](./REFERENCE.md)、[docs/internals/](../internals)

本ドキュメントは `src/lsp/` 配下に実装された VBA Language Server の機能仕様と使い方を記述します。
LSP サーバーは **別プロセスを持たないインライン拡張** として動作します（`vscode-languageserver` ライブラリは不使用）。

---

## アーキテクチャ概要

```
VS Code Extension Host (src/extension.ts)
    │
    ├── LSPServer (src/lsp/server.ts)          ← すべてのプロバイダーの中心
    │       ├── SymbolProvider                  (src/lsp/symbol-provider.ts)
    │       ├── HoverProvider                   (src/lsp/hover-provider.ts)
    │       ├── DefinitionProvider              (src/lsp/definition-provider.ts)
    │       ├── CompletionProvider              (src/lsp/completion-provider.ts)
    │       ├── ReferencesProvider              (src/lsp/references-provider.ts)
    │       ├── RenameProvider                  (src/lsp/rename-provider.ts)
    │       ├── CodeLensProvider                (src/lsp/code-lens-provider.ts)
    │       ├── FoldingRangeProvider            (src/lsp/folding-range-provider.ts)
    │       ├── TestDiscovery                   (src/lsp/test-discovery.ts)
    │       └── TestRunner                      (src/lsp/test-runner.ts)
    │
    └── VBADebugAdapterFactory                  (src/lsp/vscode-debug-adapter.ts)
            └── VBAInlineDebugAdapter
                    └── DebugAdapter            (src/lsp/debug-adapter.ts)
                            └── VBADebugSession + Worker (debug-session.ts / debug-worker.ts)
```

- **LSPServer** は `didOpen` / `didChange` / `didClose` でドキュメントを管理し、各プロバイダーの呼び出し口となる。
- 拡張機能の `activate()` はすべての VS Code プロバイダーを登録し、`LSPServer` インスタンスを共有する。
- インライン実装のため通常のLSPプロセスは起動せず、Node.js IPCも不要。DAP実行時だけ、評価器を専用Workerで動かす。
- DAP Workerは `SharedArrayBuffer` と `Atomics` で停止・再開を同期し、停止中のメッセージは `receiveMessageOnPort` で処理する。

---

## 実装済み機能一覧

| 機能 | VS Code 操作 | 実装ファイル |
|---|---|---|
| Diagnostics | 自動（保存不要） | `server.ts`, `code-lens-provider.ts`, `host-mock-advisor.ts` |
| Document Symbol | アウトラインペイン / `@` 検索 | `symbol-provider.ts` |
| Hover | カーソルホバー | `hover-provider.ts` |
| Go to Definition | F12 | `definition-provider.ts` |
| Completion | Ctrl+Space / 入力時 | `completion-provider.ts` |
| Find All References | Shift+F12 | `references-provider.ts` |
| Rename Symbol | F2 | `rename-provider.ts` |
| Code Lens | 行上部のインライン UI | `code-lens-provider.ts` |
| Folding Range | 折りたたみ / Sticky Scroll | `folding-range-provider.ts` |
| Test Discovery / Run | テストエクスプローラー（個別実行対応） | `test-discovery.ts`, `test-runner.ts` |
| DAP Debugger | F5 / デバッグビュー | `debug-adapter.ts`, `vscode-debug-adapter.ts` |

---

## 各機能の詳細

### Diagnostics（構文エラー・警告）

パースエラーと Dead code 警告を `DiagnosticCollection` に出力する。ドキュメントを開く・編集するたびに自動更新。

**出力する診断の種類:**

| 種別 | severity | 内容 |
|---|---|---|
| 構文エラー | Error (1) | `Parser` が検出した構文違反 |
| Dead code 警告 | Warning (2) | Private プロシージャで外部参照が 0 件のもの |
| VBA lint | Warning / Information | VBA001〜VBA012（ByRef、省略宣言、型・構文上の落とし穴など） |
| ホスト依存警告 | Information | `Application`、`Worksheets`、`ThisWorkbook` などのExcel/Access識別子。`__mocks__` 作成案内を含む |

**実装詳細:**
- `Parser.diagnostics` は 1-based line/column → LSP 変換時に `line - 1`, `character - 1`
- Dead code = `isPrivate && refCount === 0`（`CodeLensProvider.getDeadCodeWarnings()` と連携）

**テスト:** `tests/lsp/server-diagnostics.test.ts`, `tests/lsp/lsp-dead-code.test.ts`

`__mocks__` の追加・変更はファイル監視で再スキャンされ、対応するホスト依存診断を再計算する。診断スキャナーはJavaScriptを実行せず、CommonJSの静的なexport名だけを収集する。したがって、計算プロパティや実行時に生成されるメンバーは診断だけでは解決しない。

---

### Document Symbol（アウトライン表示）

ファイル内のシンボルをアウトラインペインに表示する。

**対応するシンボル種別:**

| VBA | VS Code SymbolKind |
|---|---|
| Sub / Function | Function (11) |
| Type | Struct (22) |
| Class / Module変数 | Variable (12) |

**実装詳細:**
- `SymbolProvider` の enum 値は 1-based。VS Code の `SymbolKind` は 0-based のため `kind - 1` で変換する。

**テスト:** `tests/lsp/lsp-document-symbol.test.ts`, `tests/lsp/lsp-symbol-provider.test.ts`

---

### Hover（ホバー情報）

カーソル位置のシンボルに関する情報をツールチップで表示する。

**表示内容:**
- プロシージャ: シグネチャ（名前・パラメーター・戻り値型）
- 変数: 宣言時の型

**テスト:** `tests/lsp/lsp-hover.test.ts`

---

### Go to Definition（定義へジャンプ）

F12 でカーソル位置のシンボルの宣言位置へジャンプする。

**実装詳細:**
- `DefinitionProvider` はシンボルテーブルを参照し、宣言の `loc` を返す。
- 組み込み関数は対象外。

**テスト:** `tests/lsp/lsp-definition.test.ts`

---

### Completion（自動補完）

入力中に候補リストを表示する。

**補完対象:**
- VBA キーワード（`Sub`, `Function`, `Dim`, `If` … 50種類以上）
- 同一モジュール内の Sub/Function 名
- 同一モジュール内の変数名

**テスト:** `tests/lsp/lsp-completion.test.ts`

---

### Find All References（参照箇所の検索）

Shift+F12 でカーソル位置のシンボルが使用されているすべての位置を列挙する。

**実装詳細 (text-based search):**
- AST ではなくソーステキストを対象に正規表現で検索する（`(?<![a-zA-Z0-9_])word(?![a-zA-Z0-9_])`、大文字小文字無視）
- 純コメント行・インラインコメント後・文字列リテラル内はスキップ
- `includeDeclaration=false` 時は宣言位置を結果から除外する

**API:**
```typescript
findAllReferences(sourceText, targetWord, uri, statements, includeDeclaration): LocationInfo[]
```

**テスト:** `tests/lsp/lsp-references.test.ts`

---

### Rename Symbol（シンボル名の変更）

F2 でシンボルをリネームし、ファイル内すべての参照を一括置換する。

**実装詳細:**
- `findAllReferences(includeDeclaration=true)` で宣言を含む全参照を取得し、`TextEdit[]` を返す。
- VS Code の `WorkspaceEdit` 経由で適用される。

**テスト:** `tests/lsp/lsp-references.test.ts`（Rename テストを含む）

---

### Code Lens（プロシージャ行上部の UI）

各 Sub/Function の宣言行の上部に最大 3 つのレンズ項目を表示する。

| レンズ | 条件 | コマンド |
|---|---|---|
| `▶ Run` | 必須パラメーターなし | `vba-runner.runProcedure` |
| `N references` | 常に表示 | `vba-runner.findReferences` |
| `⚠ 0 references` | Private かつ 0 参照 | `vba-runner.findReferences` |
| `✓ テスト済み` | `Test_*` プロシージャが参照 | `vba-runner.goToTest` |
| `未テスト` | テスト参照なし | `vba-runner.generateTest` |

**実装詳細:**
- 参照数カウント: `findAllReferences` の結果からプロシージャ本体の行範囲内を除外し外部参照のみを計算する
  ```typescript
  const externalRefs = refs.filter(r =>
      r.range.start.line < line || r.range.start.line > endLine
  );
  const refCount = externalRefs.length;
  ```
- テスト済み判定: `Test_*` プロシージャの本文内に対象名が全単語マッチで登場するか確認

#### `▶ Run` の実行仕様（`vba-runner.runProcedure`）

| 項目 | 仕様 |
|---|---|
| トリガー | Code Lens `▶ Run` クリック時のみ（オンデマンド） |
| ファイル収集 | 実行ファイルと**同ディレクトリ**の `.bas`/`.cls` を全件収集 |
| キャッシュ | なし。毎回ディスクから再読み込みする |
| `.cls` の扱い | ファイル名をクラス名として `Class ClassName … End Class` でラップして結合 |
| 実行対象 | 収集した全ファイルを結合した AST 上で対象プロシージャを呼び出す |
| 出力先 | `VBA Runner` 出力チャネル（実行後に自動表示） |

**モジュール間参照の解決例:**

```
tests/vba/
├── AssertHelper.cls        → Class AssertHelper ... End Class としてラップ
├── ArgCountTest.bas        → そのまま結合
└── ArgCountTest_runner.bas → ▶ Run を押した対象ファイル
```

3 ファイルを結合して評価するため、`Dim assert As New AssertHelper` や `ArgCountTest.Test_Foo` のようなモジュール間参照が解決される。

**制約:**
- 同ディレクトリ外のファイルは収集しない
- ファイルの収集順はOSのディレクトリ列挙順（不定）
- 引数ありプロシージャには `▶ Run` は表示されない（Code Lens 側の制約）

**テスト:** `tests/lsp/lsp-code-lens.test.ts`

---

### Folding Range（折りたたみ / Sticky Scroll）

VBA のブロック構造を VS Code の折りたたみ機能と Sticky Scroll に対応させる。

**対応するブロック:**

| VBA 構文 | 説明 |
|---|---|
| `Sub` / `Function` / `Property` | プロシージャ本体 |
| `For ... Next` / `For Each ... Next` | For ループ |
| `Do ... Loop` | Do ループ |
| `While ... Wend` | While ループ |
| `With ... End With` | With ブロック |
| `If ... End If` | If ブロック（単行 If を除く） |
| `Select Case ... End Select` | Select Case ブロック |
| `Type ... End Type` | ユーザー定義型 |
| `Enum ... End Enum` | 列挙型 |

**Sticky Scroll の動作:**

エディター上部に現在カーソルが属するブロックのヘッダー行が固定表示される。たとえばループ本体が画面外にスクロールした場合でも `For i = 1 To 100` の行が上部に残る。

```
Sub ProcessData()           ← 常に上部に表示
    For i = 1 To 100        ← スクロールアウトしても上部に表示
        │
        │  ← カーソル位置
```

**実装詳細:**
- `FoldingRangeProvider.getFoldingRanges()` が AST を再帰的に走査して `{ startLine, endLine }` を返す（0-based）
- `IfStatement.alternate` が `IfStatement`（ElseIf チェーン）の場合は再帰的に処理する
- パーサーの `loc` フィールド（1-based）を 0-based に変換して返す

---

### Test Discovery / Run（テスト探索と実行）

`Test_` で始まる Sub を自動検出し、テストエクスプローラーに表示・実行する。

- スイート実行は `runTests`、選択項目の実行は `runTestWithEvaluation` を使い、選択したSubだけを評価する。
- 実行対象と同じディレクトリの `.bas` / `.cls` を収集し、クラスモジュールはクラス定義として結合する。
- テスト実行のキャンセルは `TestRunRequest.token` を確認し、未開始の項目を実行しない。

**テスト:** `tests/lsp/lsp-test-discovery.test.ts`, `tests/lsp/lsp-test-runner.test.ts`

---

### DAP Debugger（デバッグアダプター）

F5 で VBA コードをステップ実行する。別プロセス不要のインライン実装。

**対応 DAP コマンド:**

| コマンド | 動作 |
|---|---|
| `initialize` | デバッガー初期化、`initialized` イベント送出 |
| `launch` | 実行開始、`stopped(entry)` イベント送出 |
| `configurationDone` | DAPハンドシェイクを完了（実行開始は `launch`） |
| `setBreakpoints` | ブレークポイント登録（`condition` にVBA式、`hitCondition` に到達回数を指定可能） |
| `threads` | スレッド一覧（1スレッド固定） |
| `stackTrace` | コールスタック |
| `scopes` | スコープ一覧（Locals） |
| `variables` | 変数一覧 |
| `setVariable` | 停止中に既存のローカル変数をVBA式で評価した値へ変更（未知の名前は拒否） |
| `evaluate` | 停止中の現在スコープでVBA式を評価（ウォッチ式） |
| `continue` | 実行継続。条件を満たすブレークポイントで停止し、なければ終了 |
| `pause` | 実行中の次のステートメントで `stopped(pause)` 送出 |
| `next` | ステップオーバー、`stopped(step)` 送出 |
| `stepIn` / `stepOut` | ステップイン/アウト |
| 実行時エラー | `stopped(exception)` として発生行・スタックを通知し、実行時エラーを出力 |
| `disconnect` | 終了、`terminated` イベント送出 |

**起動設定（`.vscode/launch.json`）:**
```json
{
    "type": "vba",
    "request": "launch",
    "name": "Run VBA",
    "program": "${file}"
}
```

**実装ファイル:**
- `src/lsp/debug-adapter.ts` — DAP ロジック本体
- `src/lsp/vscode-debug-adapter.ts` — VS Code `DebugAdapterDescriptorFactory` / `DebugAdapter` ラッパー
- `src/lsp/debug-session.ts` — Worker管理、停止状態、式評価、変数変更
- `src/lsp/debug-worker.ts` — DebugHook、ステップ制御、条件・ヒット回数判定、例外位置の収集

`evaluate` は停止中のWorkerへ式を転送し、現在の環境で評価した値とVBA型名を返す。
DAPの評価コンテキストを明確にするため、入力式は括弧で囲んで評価される。

`setVariable` は既存のローカル変数だけを対象とし、入力値をVBA式として評価してからEvaluatorの型強制を通して代入する。未知の名前、実行中の変更、配列要素やオブジェクトメンバーの編集は対象外である。

条件付きブレークポイントの条件式が評価できない場合は、その到達では停止しない。`hitCondition` は正の整数として解釈し、同じ行への到達回数が閾値に達した時点で停止する。

実行時エラーは、直前のステートメントの行・スタック・変数を `stopped(exception)` として通知した後、実行終了イベントを送出する。DAPクライアントは停止イベント受信後に `stackTrace` / `variables` を取得できる。

**テスト:** `tests/lsp/lsp-debug-adapter.test.ts`, `tests/lsp/lsp-debug-session.test.ts`（15項目）、`tests/lsp/lsp-debug-hook.test.ts`

---

## テストの実行方法

```bash
# 個別テスト（例: Code Lens）
npx tsx tests/lsp/lsp-code-lens.test.ts

# 全 LSP テスト一括実行
for f in tests/lsp/*.test.ts; do
    npx tsx "$f"
done
```

---

## 今後の拡張候補

- DAPの配列要素・オブジェクトメンバー編集
- キーワード引数に対応したSignature Helpのアクティブ引数選択
- Excel/COM型ライブラリからのSignature Help定義ロード
- コールグラフの可視化
- npm パッケージ化（エンジン本体を `@vba-runner/engine` として公開）

詳細は `docs/todo/TODO.md` を参照。
