# VBA Language Server Protocol (LSP) 開発計画

既存のVBAコンパイラの `Lexer` および `Parser` を活用し、VSCode等のエディタで動作する **VBA Language Server** を構築するためのロードマップと技術設計です。

## 1. 目標と提供する機能

標準的なLSPの仕様（`vscode-languageserver`）に則り、以下の機能の提供を目指します。

1. **Diagnostics（構文エラーのリアルタイム表示）**
   - タイピング中に構文エラーや未定義変数の警告をエディタ上に波線で表示する。
2. **Document Symbol（アウトライン表示）**
   - ファイル内の `Sub`, `Function`, `Type`, モジュールレベル変数などを一覧表示し、アウトラインペインやブレッドクラムに対応する。
3. **Hover（ホバー情報表示）**
   - 変数や関数にカーソルを合わせた際、そのスコープや組み込み関数のシグネチャをツールチップで表示する。
4. **Go to Definition（定義へジャンプ）**
   - 関数呼び出しや変数参照から、その定義位置へジャンプ（F12）する機能。
5. **Completion（自動補完）**
   - `Public`, `Sub`, などの予約語の補完。
   - 同一モジュール内の変数やプロシージャ名の補完。

## 2. アーキテクチャ設計

### ディレクトリ構成案
モノレポ構成とし、既存のコンパイラ基盤を共有します。

```text
vba-runner/
├── src/
│   └── compiler/          # 既存のLexer, Parser, Evaluator
├── lsp/
│   ├── client/            # VSCode拡張機能（フロントエンド）
│   │   └── src/extension.ts
│   └── server/            # LSPサーバー（バックエンド）
│       └── src/server.ts
└── package.json
```

### 依存パッケージ
- `vscode-languageserver`: サーバーサイド用モジュール
- `vscode-languageserver-textdocument`: テキスト同期・位置情報管理
- `vscode-languageclient`: VSCode拡張用クライアント

## 3. 実装フェーズ

### Phase 1: 基礎構築とVSCode拡張のセットアップ
- `lsp/client` および `lsp/server` の初期化。
- クライアント・サーバー間のプロセス間通信（IPC）の確立。
- エディタでドキュメントを開いた際、または編集した際にサーバーへテキスト内容を送信するイベント連携の構築。

### Phase 2: 既存コンパイラのLSP向け改修（重要課題）
LSPとして動作させるためには、既存のコンパイラ基盤に以下の改修が必要です。

1. **トークン・ASTへの完全な位置情報の付与**
   - 現在 `Lexer` は `line` (行番号) のみを保持しています。LSPでは「何行目の何文字目から何文字目まで」が必須となるため、各トークンおよびASTノードに `{ line, character }` 形式の `start` と `end` 位置を追加する改修を行います。
2. **Tolerant Parsing (寛容なパース・エラー耐性)**
   - 現在の `Parser` は構文エラーに遭遇すると `throw new Error` で直ちにパースを停止します。LSPではエラー箇所をスキップして後続のコードをパースし続ける仕組み（Error Recovery）が必要です。エラーを配列に収集し、不完全なASTでも返却できるように `Parser` を拡張します。

### Phase 3: Diagnostics（エラー解析）の統合
- クライアントから受け取ったテキストを `Lexer` -> `Parser` に通す。
- Phase 2 で収集したエラーの配列（行番号、列、メッセージ）を `Diagnostic` オブジェクトに変換し、クライアントに送信して波線を表示させる。

### Phase 4: シンボルとホバー機能 (Language Features) の実装
- パースされたASTを走査（Visitorパターン等）し、変数やプロシージャの「シンボルテーブル」を構築。
- エディタからの `textDocument/documentSymbol` リクエストに対し、シンボルの一覧を返す。
- `textDocument/hover` リクエストに対し、カーソル位置のノードを特定し、関連する情報を返す。

### Phase 5: リリースとパッケージング
- `.vsix` ファイルへのパッケージング（`vsce` ツールを使用）。
- 拡張機能のインストールと利用手順のドキュメント化。

## 4. プロジェクト上の次のステップ（提案）

もしこの計画で進める場合、最初のタスクは以下のようになります。

1. **Step 1:** `Lexer` を改修し、行番号に加えて**列番号 (column)** と**トークン長 (length)** を取得できるようにする。
2. **Step 2:** `Parser` を改修し、ASTノード（`Statement` や `Expression`）に `start` と `end` の位置情報を付与する。
3. **Step 3:** LSPの初期パッケージ構成（`package.json`の更新など）を作成し、Hello WorldのDiagnosticsを出力するサーバーを立ち上げる。
