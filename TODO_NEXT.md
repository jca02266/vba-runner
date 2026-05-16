# 次期開発ロードマップ (Phase 2: IDE統合と開発エコシステム)

MS-VBAL 仕様書で個別セクションに名前の付いた構文要素・標準ライブラリ関数（Phase 1）の実装が完了したことを受け、本プロジェクトの最終ゴールである**「VSCode内で完結するモダンなVBA統合開発環境（IDE）」**の構築に向けた次期TODOリストです。

> なお、Phase 1 はリスト化された機能の網羅であり、仕様書本文に記載された **ランタイム挙動の細部** までは未確認です。これらは [`TODO.md` の「VBA ランタイム挙動」](TODO.md#vba-ランタイム挙動) で個別にトラッキングしています。

## 1. エディタ支援（Language Server Protocol: LSP）
VSCode拡張機能として動作し、リアルタイムなコード解析とコーディング支援を提供します。詳細は `LSP.md` を参照。

- [ ] **Tolerant Parsing（エラー耐性）の実装**
  - [ ] Lexer: トークンに列番号（`column`）と長さ（`length`）を付与する。
  - [ ] Parser: ASTノードに `start` と `end` の厳密な位置情報を付与する。
  - [ ] Parser: エラー発生時にパースを中断せず、エラーリストとして収集する仕組み（Error Recovery）の実装。
- [ ] **Diagnostics（構文エラー表示）**
  - [ ] 解析されたエラーリストを波線としてエディタ上にフィードバックする。
- [ ] **シンボル解決とホバー機能**
  - [ ] `textDocument/documentSymbol` の実装（アウトライン表示）。
  - [ ] `textDocument/hover` の実装（変数や関数の型情報ツールチップ）。
  - [ ] `textDocument/definition` の実装（F12による定義元ジャンプ）。
- [ ] **自動補完（Autocomplete）**
  - [ ] `textDocument/completion` の実装（標準関数、モジュール内変数の補完）。

## 2. テストエクスプローラーの統合 (VSCode Testing API)
VBAのテストコードをVSCode上でシームレスに実行・確認できる環境を構築します。

- [ ] **テストの自動検出**
  - [ ] ワークスペース内のVBAファイルから、テスト用プロシージャ（例: `Test_` から始まるSub）を抽出し、VSCodeのテストエクスプローラーに一覧表示する。
- [ ] **エディタからの直接実行 (CodeLens)**
  - [ ] テスト定義の上に `▶ Run Test` ボタンを表示。
  - [ ] 実行結果（Pass/Fail）およびエラーメッセージをVSCodeのTesting UIに直接反映させる。
- [ ] **Test Double APIの拡張**
  - [ ] テストランナー上で `MsgBox` や `Shell` などの副作用を検証できるSpy/Mock APIの提供。

## 3. インタラクティブな実行環境とデバッグ (DAP)
Excelを持たない環境（Mac/Linux/Web）でも、VBAマクロを実行・デバッグ可能にします。

- [ ] **標準入出力のVSCode連携**
  - [ ] `Debug.Print` の出力をVSCodeの Output Panel または内蔵ターミナルにルーティング。
  - [ ] `MsgBox` や `InputBox` の呼び出しを、VSCodeの標準UI（`vscode.window`）にバインドして表示。
- [ ] **Debug Adapter Protocol (DAP) の基礎構築**
  - [ ] Evaluatorにステップ実行（Step Over, Step Into）のフックを追加。
  - [ ] エディタで設定されたブレークポイント（赤い丸）で実行を一時停止する機能。
  - [ ] 停止時のスコープ内の変数（ローカル変数など）をVSCodeのデバッグパネルに一覧表示する機能。

## 4. Web拡張機能としてのパッケージングとリリース
ローカル環境に依存しない、ブラウザ完結の開発体験を提供します。

- [ ] **Web拡張機能 (Web Extension) 化**
  - [ ] Node.jsのファイルシステム（`fs`）等に依存する処理を抽象化し、`vscode.workspace.fs` に置き換える。
  - [ ] `esbuild` や `webpack` でブラウザ用（VSCode for the Web, github.dev 向け）にバンドル。
- [ ] **VSCode Marketplace への公開**
  - [ ] 拡張機能のアイコン、README の整備。
  - [ ] `.vsix` ファイルのビルドパイプラインの構築と公開手続き。
## 5. テスト基盤のエンタープライズ強化 (Advanced Testing API)
実際の業務マクロを効率よくテストするための基盤拡張です。

- [ ] **副作用の検証機能 (Spy / Mocking API)**
  - [ ] `VBATest.spy("MsgBox")` のように呼び出し回数や引数をアサーション可能にする。
- [ ] **日時や環境の仮想化 (Time Mocking)**
  - [ ] `vbaTest.mockDate('2024-12-31')` などで `Now` や `Date` の結果をテストごとに固定する。
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
- [x] `path-browserify` によるブラウザ互換のパス操作。
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
