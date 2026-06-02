# VBA Runner — VBA実行環境 + リファクタリング支援ツール

Excel 不要で VBA コードを実行・テスト・静的解析できる実行環境です。

## VBA Runnerの目的

- **Excel 不要での実行・確認**: Node.js 環境で直接 VBA の構文とロジックを実行・検証できます。
- **リファクタリングの支援**: `vba-analyzer` を使ってVBAコードベースの問題箇所を特定し、純粋な関数を安全に切り出せます。
- **ユニットテストの実行**: TypeScript のテストランナーを通じて、VBA 関数にモックデータとアサーションを与えてテストを自動化できます。

VS Code などのモダンな環境で、AIの支援を受けながら VBA ソースのリファクタリングとテストを進めることができます。

## 提供パッケージ

| パッケージ | 内容 | 詳細 |
|---|---|---|
| `vba-runner`（npm ライブラリ） | VBA実行エンジン + テストランナー + CLI ツール | [build/runner/README.md](build/runner/README.md) |
| VBA Runner（VS Code 拡張機能） | LSP 統合・デバッガー・コール階層表示 | [build/extension/README.md](build/extension/README.md) |
| Playground（Web UI デモ） | ブラウザ上でVBAを即試せるデモサイト | [build/playground/README.md](build/playground/README.md) |

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
├── test-libs/           VBARunnerクラス、CLI ツール（vba-analyzer 等）
├── tests/spec/          VBA言語仕様テスト群
├── sample/              リファクタリング事例（TaskScheduler）
│   ├── src/vba/         プロダクションコード（VBA）
│   └── tests/ts/        サンプルのユニットテスト（TypeScript）
├── build/
│   ├── runner/          npm パッケージ配布物
│   ├── extension/       VS Code 拡張機能配布物（.vsix）
│   └── playground/      Web UI 配布物（Vite）
└── docs/                詳細設計ガイド・実践例
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

| ドキュメント | 内容 |
|---|---|
| [REFERENCE.md](REFERENCE.md) | 詳細仕様（サンドボックス方針・型システム・モック・拡張機能など） |
| [TODO.md](TODO.md) | 次期開発ロードマップ（Phase 2: VS Code IDE統合） |
| [TODO_SPEC.md](TODO_SPEC.md) | MS-VBAL仕様書準拠の実装進捗 |
| [FOR_AI.md](FOR_AI.md) | AIによるリファクタリング支援のガイド |
| [LSP.md](LSP.md) | VS Code LSP の設計・実装方針 |
| [docs/](docs/) | テスト設計・リファクタリング手法・モックガイド |
