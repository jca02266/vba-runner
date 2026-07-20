# VBA Runner — VBA実行環境 + リファクタリング支援ツール

Excel 不要で VBA コードを実行・テスト・静的解析できる実行環境です。
AI と組み合わせることで、リファクタリング提案を「仮説」のまま終わらせず、**その場で実行して検証済み**にできます。

> English: [README.md](./README.md)

## VBA Runner の付加価値

AI は VBA コードを読んでリファクタリング案を提案できますが、VBA Runner がない状態では結果の確認に Excel を開く必要があり、イテレーションが重くなります。

VBA Runner は、次の問題を解決できます。

| 課題 | VBA Runner なし | VBA Runner あり |
|---|---|---|
| **提案の検証** | Excel で手動確認（重い） | その場で実行して即検証 |
| **大規模コード把握** | トークン上限で全量渡せない | `vba-analyzer` のアウトラインを AI に渡す |
| **リファクタリングの安全網** | なし（勘と経験） | 変更前に実行 → スナップショットテストを自動生成 |
| **反復速度** | 変更 → Excel 貼り付け → 手動実行 | 変更 → テスト実行（数秒） |

## 何をしたいですか？

| 目的 | はじめる場所 |
|---|---|
| VBA を TypeScript から実行・テストする | [vba-runner](build/runner/README.ja.md) |
| Excel ファイルから VBA を抽出・書き戻す | [vba-extractor](build/extractor/README.ja.md) |
| VS Code で VBA を編集・確認する | [VS Code 拡張機能](build/extension/README.ja.md) |
| ブラウザで VBA を試す | [VBA Web Runner](build/playground/README.md) |
| AI とレガシー VBA を改善する | [実践チュートリアル](docs/TUTORIAL.md) |
| すべてのガイドから目的に合うものを探す | [ドキュメントガイド](docs/README.md) |
| このリポジトリを開発する | [開発者ガイド](CONTRIBUTING.md) |

## 提供パッケージ

| パッケージ | 内容 | 正本 |
|---|---|---|
| `vba-runner`（npm ライブラリ） | VBA 実行エンジン + テストランナー + CLI | [build/runner/README.ja.md](build/runner/README.ja.md) |
| `vba-extractor`（npm CLI） | Office ファイルから VBA ソースを export / import | [build/extractor/README.ja.md](build/extractor/README.ja.md) |
| VBA Runner（VS Code 拡張機能） | LSP・デバッガー・コール階層表示 | [build/extension/README.ja.md](build/extension/README.ja.md) |
| VBA Web Runner（Web UI デモ） | ブラウザ上で VBA を即試せるデモ | [build/playground/README.md](build/playground/README.md) |

パッケージのインストール手順や API は、各正本を参照してください。
リポジトリを clone して本体を開発する場合は [CONTRIBUTING.md](CONTRIBUTING.md) を参照してください。

## 関連ドキュメント

| 文書 | 役割 |
|---|---|
| [docs/README.md](docs/README.md) | 目的別の共通ハブ |
| [FOR_AI.md](FOR_AI.md) | 利用者の VBA を AI が扱うための指示書 |
| [REFERENCE.md](REFERENCE.md) | 詳細仕様（Sandbox・型・モックなど） |
| [CONTRIBUTING.md](CONTRIBUTING.md) | リポジトリ開発者向け入口 |
