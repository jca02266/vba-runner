# 次期開発ロードマップ (Phase 2: IDE統合と開発エコシステム)

MS-VBAL 仕様書で個別セクションに名前の付いた構文要素・標準ライブラリ関数（Phase 1）の実装が完了したことを受け、本プロジェクトの最終ゴールである**「VSCode内で完結するモダンなVBA統合開発環境（IDE）」**の構築に向けた次期TODOリストです。

> なお、Phase 1 はリスト化された機能の網羅であり、仕様書本文に記載された **ランタイム挙動の細部** までは未確認です。これらは [`TODO.md` の「VBA ランタイム挙動」](TODO.md#vba-ランタイム挙動) で個別にトラッキングしています。

## 1. エディタ支援（Language Server Protocol: LSP）
⚠️ **部分完了** — Hover / Definition / Completion は VSCode 拡張機能として動作中。Diagnostics と DocumentSymbol は実装済みだが extension.ts 未接続。

- [x] **Tolerant Parsing（エラー耐性）**
  - [x] Lexer: 列番号（`column`）を全トークンに付与。テスト: `lexer-column.test.ts`
  - [x] Parser: ASTノードに `loc: {start, end}` の位置情報を付与。テスト: `parser-error-recovery.test.ts`
  - [x] Error Recovery: エラー発生時にパースを中断せず収集。
- [ ] **Diagnostics（構文エラー表示）**
  - [x] Parser が `ast.diagnostics[]` にエラーを収集（エラー耐性パース）。テスト: `lsp-diagnostics.test.ts`
  - [ ] DiagnosticsProvider クラスの作成（`ast.diagnostics` → `vscode.Diagnostic[]` 変換）
  - [ ] extension.ts への登録（`createDiagnosticCollection`）
- [x] **シンボル解決とホバー機能**
  - [ ] `textDocument/documentSymbol`（アウトライン表示）— SymbolProvider 実装済み（テスト: `lsp-symbol-provider.test.ts`）だが extension.ts 未接続
  - [x] `textDocument/hover`（型情報ツールチップ）。テスト: `lsp-hover.test.ts`
  - [x] `textDocument/definition`（F12による定義元ジャンプ）。テスト: `lsp-definition.test.ts`
- [x] **自動補完（Autocomplete）**
  - [x] `textDocument/completion`（標準関数、変数補完）。テスト: `lsp-completion.test.ts`

## 2. テストエクスプローラーの統合 (VSCode Testing API)
✅ **完了** — VBAのテストコードをVSCode上でシームレスに実行・確認。

- [x] **テストの自動検出**
  - [x] `Test_` で始まるプロシージャを抽出、VSCodeのテストエクスプローラーに表示。テスト: `lsp-test-discovery.test.ts`
- [x] **エディタからの直接実行**
  - [x] Pass/Fail 結果および エラーメッセージをTesting UIに反映。テスト: `lsp-test-runner.test.ts`
- [x] **Test Double API**
  - [x] Spy/Mock API で `MsgBox`, `Shell` などの副作用を検証。テスト: `spy-mock-api.test.ts`

## 3. インタラクティブな実行環境とデバッグ (DAP)
⚠️ **部分完了** — DebugAdapter / Debugger の実装とテストは存在するが、extension.ts への統合は未実施。

- [x] **標準入出力のVSCode連携**
  - [x] `Debug.Print` → Output Panel（console.log で実装）
  - [ ] `MsgBox` / `InputBox` → VSCode標準UI（現状はスタブ動作のみ）
- [ ] **Debug Adapter Protocol (DAP) — extension.ts 未接続**
  - [x] ステップ実行フック（Step Over, Into）実装済み。テスト: `lsp-debugger.test.ts`
  - [x] ブレークポイント実行一時停止 実装済み。テスト: `lsp-debugger.test.ts`
  - [x] デバッグパネルに変数一覧表示 実装済み。テスト: `lsp-debug-adapter.test.ts`
  - [ ] `vscode.debug.registerDebugAdapterDescriptorFactory` による VSCode への登録

## 4. パッケージングとリリース
- [x] **Extension インフラストラクチャ**
  - [x] `src/extension.ts`: Node.js版 (Desktop用)
  - [x] `language-configuration.json`: 言語設定
  - [x] `syntaxes/vba.tmLanguage.json`: TextMate グラマー
  - [x] `package.json`: VSCode extension metadata
  - [x] `.vscodeignore`: パッケージ除外ルール
- [ ] **Web拡張機能 (Web Extension) 化** — 取りやめ中
  - [ ] `src/extension-web.ts`: Browser版 (vscode.dev, github.dev 向け)
  - [ ] `esbuild` でブラウザ用にバンドル（evaluator が Node.js `path` に依存するため要先行対応）
- [ ] **VSCode Marketplace への公開**
  - [ ] 拡張機能のアイコン、README の整備。
  - [ ] `.vsix` ファイルのビルドパイプラインの検証。
  - [ ] VSCode Marketplace への公開手続き。
## 5. テスト基盤のエンタープライズ強化 (Advanced Testing API)
実際の業務マクロを効率よくテストするための基盤拡張です。

- [x] **副作用の検証機能 (Spy / Mocking API)**
  - [x] `VBATest.spy("MsgBox")` / `Evaluator.spy(name, returnFn?)` で呼び出し回数・引数を記録。
  - [x] `SpyRecord`: `callCount`, `calls`, `lastCall`, `returnValues`, `calledWith()`, `reset()` を提供。
  - [x] `returnFn` 指定で戻り値をオーバーライド可能（例: `MsgBox` を常に `vbYes=6` にする）。
  - テスト: `spy-mock-api.test.ts`
- [x] **日時や環境の仮想化 (Time Mocking)**
  - [x] `vbaTest.mockDate('2024-12-31')` などで `Now` / `Date` / `Time` / `Timer` の結果をテストごとに固定する。
  - [x] `Evaluator.setNowFn(fn)` で任意の Date ファクトリを注入。`null` 渡しで実時刻に復元。
  - テスト: `time-mocking.test.ts`
- [x] **インメモリ・ファイルシステム (In-Memory FS)**
  - [x] `FileSystem` インターフェースによる抽象化と `MemoryFileSystem` の基礎実装。
  - [ ] サンドボックスのファイル操作をメモリ上に仮想化し、実ディスクI/Oなしで並列テストを高速化する。
  - [ ] **VFS と Node.js FS の相互互換性向上**: VFS モードで `./sandbox` や `../foo/bar/sandbox` のような相対パスを受け入れても、内部的には `/sandbox` を起点にルートディレクトリとして解決する仕組みの実装。これにより、同じテストコード（相対パスを使用）が VFS と実ファイルシステムの両モードで同一に動作するようになる。

## 6. レガシー移行支援機能 (Migration Tooling)
VBAを実行するだけでなく、モダンな言語環境への移行を支援するツール群です。

- [ ] **VBA to TypeScript トランスパイラ**
  - [ ] 構築済みのASTから、TypeScriptのコードを自動生成する Generator クラスの実装。
- [x] **デバッグとスタックトレースの精緻化（部分実装）**
  - [x] ASTノードに `line?: number` を付与。`parseStatement` ラッパーで全ステートメントに自動付与。
  - [x] `Evaluator.currentLine` で実行中の行番号を追跡。`throwVbaError` がエラーメッセージに `(line N)` と `err.vbaLine` を付与。
  - [ ] コールスタック（呼び出し元の行番号）の出力。`Erl` 関数のサポート。

## 7. Webコンパイラ (Web UI) の改善課題
ブラウザ上で完結する実行環境としての完成度を高めます。
## 完了済み
- [x] VFS (仮想ファイルシステム) インターフェースと `MemoryFileSystem` の実装。
- [x] Node.js `fs` モジュールの切り離しと DI (依存性注入) 化。
- [x] `Kill` ステートメントのサポート（VBA互換の実行時エラー 53 対応）。
- [x] ビルドを妨げていた AST インターフェースの不整合 (`scope`, `objectType`) の修正。
- [x] 本番環境向けの `npm run build` の成功確認。

## 今後の課題
### VFS & ファイル操作
- [ ] `Kill` ステートメントのワイルドカード対応 (例: `Kill "*.txt"`)。
- [ ] `Dir` 関数の完全な実装 (ディレクトリ内の列挙)。
- [ ] `MemoryFileSystem` の永続化 (localStorage/IndexedDB との統合)。
- [ ] 未実装のファイル操作ステートメントの実装: `Open`, `Close`, `Print #`, `Line Input #`, `Get`, `Put`。

### Web UI の強化
- [ ] **仮想ファイルエクスプローラー**: `MemoryFileSystem` 内のファイルを管理・閲覧できるサイドバーの追加。
- [ ] **エラーコンソール**: 実行時エラーやコンパイルエラーを、行番号付きでより分かりやすく表示。
- [ ] **エディタの同期**: 現在編集中のファイルが正しくエディタに反映される仕組みの構築。
- [ ] **Sandbox ファイル一覧**: Sandbox 内に存在するファイルをリアルタイムで表示するパネル。

### 実行エンジンの改善
- [ ] `VbaDate` および `VbaCurrency` の完全な実装。
- [ ] エバリュエーターにおける `Option Base` (配列の添字開始番号) のサポート。
- [ ] 標準ライブラリオブジェクトのスタブ追加 (`Scripting.Dictionary` など)。

---
**最終ビジョン**:  
VBA開発者がVBE（Visual Basic Editor）を開くことなく、VSCode（またはブラウザ）上で、LSPによるモダンな補完を受けながらコードを書き、Testing APIで単体テストを即座に回し、必要に応じてデバッガでステップ実行できる「世界一快適なVBA開発環境」を実現する。
