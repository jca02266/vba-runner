# VBA Runner — VBA実行環境 + リファクタリング支援ツール

Excel 不要で VBA コードを実行・テスト・静的解析できる実行環境です。
AI と組み合わせることで、リファクタリング提案を「仮説」のまま終わらせず、**その場で実行して検証済み**にできます。

## VBA Runnerの付加価値

AI は VBA コードを読んでリファクタリング案を提案できますが、VBA Runner がない状態では結果の確認に Excel を開く必要があり、イテレーションが重くなります。

| 課題 | VBA Runner なし | VBA Runner あり |
|---|---|---|
| **提案の検証** | Excel で手動確認（重い） | その場で実行して即検証 |
| **大規模コード把握** | トークン上限で全量渡せない | `vba-analyzer` のアウトラインをAIに渡す |
| **リファクタリングの安全網** | なし（勘と経験） | 変更前に実行 → スナップショットテストを自動生成 |
| **反復速度** | 変更 → Excel貼り付け → 手動実行 | 変更 → テスト実行（数秒） |

## 提供パッケージ

| パッケージ | 内容 | 詳細 |
|---|---|---|
| `vba-runner`（npm ライブラリ） | VBA実行エンジン + テストランナー + CLI ツール | [build/runner/README.md](build/runner/README.md) |
| `vba-extractor`（npm CLI） | Excel 等の Office ファイルから VBA ソースを export/import | [build/extractor/README.md](build/extractor/README.md) |
| VBA Runner（VS Code 拡張機能） | LSP 統合・デバッガー・コール階層表示 | [build/extension/README.md](build/extension/README.md) |
| VBA Web Runner（Web UI デモ） | ブラウザ上でVBAを即試せるデモサイト | [build/playground/README.md](build/playground/README.md) |

## 用途別の始め方

| やりたいこと | 最初に読むドキュメント |
|---|---|
| VBA 関数を TypeScript からテストしたい | [build/runner/README.md](build/runner/README.md) |
| Excel の VBA ソースをファイルに書き出したい | [build/extractor/README.md](build/extractor/README.md) |
| レガシー VBA のリファクタリングを AI に依頼したい | [FOR_AI.md](FOR_AI.md) |
| 抽出 → VS Code で確認 → AI でリファクタリング/機能追加 → 書き戻しの一連の流れを体験したい | [docs/TUTORIAL.md](docs/TUTORIAL.md) |
| 同じ流れを、コマンドの詳細を知らずに AI への指示だけで進めたい | [docs/TUTORIAL_AI.md](docs/TUTORIAL_AI.md) |
| 自分でリファクタリング手法を学びたい | [docs/REFACTORING_GUIDE.md](docs/REFACTORING_GUIDE.md) → [docs/INDEX.md](docs/INDEX.md) |
| VS Code で VBA を快適に編集したい | [build/extension/README.md](build/extension/README.md) |

## クイックスタート

```bash
# 1. リポジトリをクローン
git clone https://github.com/jca02266/vba-runner.git
cd vba-runner

# 2. 依存関係のインストール
npm install

# 3. サンプルのユニットテストを実行して動作確認
npx tsx sample/tests/ts/TaskScheduler_Core.test.ts

# 4. 仕様テストを一括実行
npm test
```

リファクタリング対象の VBA ソースを `vba-runner/` と同じ階層に配置し、テストを書き始めることができます。

## ディレクトリ構成

```
vba-runner/
├── src/engine/          VBA用 Lexer / Parser / Evaluator（コアエンジン）
├── src/extension.ts     VS Code 拡張機能のエントリポイント
├── src/lsp/             LSP 実装
├── src/App.tsx          Web UI（React）
├── test-libs/           VBARunner クラス・CLI ツール（vba-analyzer 等）の実装
├── tests/spec/          VBA言語仕様テスト群
├── tests/vba/           実VBAと動作比較するためのテスト
├── sample/              リファクタリング事例（TaskScheduler）
│   ├── src/vba/         プロダクションコード（VBA）
│   └── tests/ts/        サンプルのユニットテスト（TypeScript）
├── tools/extractor/     vba-extractor CLI の実装
├── build/
│   ├── runner/          npm パッケージ配布物
│   ├── extractor/       vba-extractor npm パッケージ配布物
│   ├── extension/       VS Code 拡張機能配布物（.vsix）
│   └── playground/      VBA Web Runner 配布物（Vite）
└── docs/                チュートリアル・詳細設計ガイド・実践例（14本）
```

VBA ソースとテストを並べて配置する場合の例:

```
project-dir/
├── vba-runner/          git clone で配置
├── massive-vba-project/ リファクタリングしたい VBA ソース
│   ├── Module1.bas
│   └── Module2.bas
└── test/                TypeScript で書いたテストコード
    ├── Module1Test/test.ts
    └── Module2Test/test.ts
```

## ドキュメント

| ドキュメント | 対象 | 内容 |
|---|---|---|
| [REFERENCE.md](REFERENCE.md) | 利用者・開発者 | 詳細仕様（サンドボックス方針・型システム・モック・拡張機能など） |
| [FOR_AI.md](FOR_AI.md) | VBAリファクタリングをAIに依頼する場合 | AI向け操作手順・リファクタリングサイクル（Phase 1-7） |
| [LSP.md](LSP.md) | VS Code拡張開発者 | LSP の設計・実装仕様 |
| [docs/](docs/) | VBAテスト・リファクタリング実践者 | 一連の流れを学べるチュートリアル・テスト設計原則・リファクタリング手法・モックガイド（14本） |
| [INDEX.md](INDEX.md) | — | 全ドキュメント一覧・評価・再編提案 |
