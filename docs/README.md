# VBA Runner Documentation

目的から探す共通ハブです。パッケージの使い方そのものは各配布物の README が正本です。

| いま知りたいこと | 対象 | 先に読む場所 |
|---|---|---|
| プロジェクト全体の入口 | 全員 | [ルート README](../README.ja.md) |
| npm / CLI / VS Code 拡張の使い方 | 利用者 | 下の「利用者向け: はじめる」 |
| テスト・リファクタリングの実践 | 利用者 | 下の「利用者向け: AI とレガシー VBA を改善する」 |
| CLI の npm / clone 対応表 | 全員 | [CONTRIBUTING.md](../CONTRIBUTING.md#cli-コマンド対応表) |
| このリポジトリ自体の開発 | 開発者 | [CONTRIBUTING.md](../CONTRIBUTING.md) |

詳細な読み順・問題別検索が必要なときだけ [INDEX.md](INDEX.md)（任意）を参照してください。

---

## 利用者向け: はじめる

| 目的 | 正本 |
|---|---|
| TypeScript から VBA を実行・テストする | [vba-runner](../build/runner/README.ja.md) |
| Excel ファイルの VBA を取り出す／書き戻す | [vba-extractor](../build/extractor/README.ja.md) |
| VS Code で編集・確認する | [VS Code 拡張機能](../build/extension/README.ja.md) |
| ブラウザで試す | [VBA Web Runner](../build/playground/README.md) |

---

## 利用者向け: AI とレガシー VBA を改善する

| 目的 | 文書 |
|---|---|
| 最初の一連の体験（抽出 → 編集 → 書き戻し） | [TUTORIAL.md](TUTORIAL.md) |
| AI への指示だけで同じ流れを進める | [TUTORIAL_AI.md](TUTORIAL_AI.md) |
| AI に作業を任せるときの実行契約 | [FOR_AI.md](../FOR_AI.md) |
| テスト可能化の原則 | [TESTING_STRATEGY.md](TESTING_STRATEGY.md) |
| Domain Logic と I/O の分離 | [REFACTORING_GUIDE.md](REFACTORING_GUIDE.md) |
| JavaScript から VBA をテストする | [TEST_FRAMEWORK_GUIDE.md](TEST_FRAMEWORK_GUIDE.md) |
| Excel オブジェクトをモックする | [MOCK_GUIDE.md](MOCK_GUIDE.md) |

### 手法・事例（必要なとき）

| 目的 | 文書 |
|---|---|
| 手法名のカタログ | [REFACTORING_TESTING_CATALOG.md](REFACTORING_TESTING_CATALOG.md) |
| TaskScheduler のリファクタリング実例 | [REFACTORING_EXAMPLE.md](REFACTORING_EXAMPLE.md) |
| 統合テストの実例 | [INTEGRATION_TEST_EXAMPLE.md](INTEGRATION_TEST_EXAMPLE.md) |
| テーブル駆動リファクタリング | [TABLE_DRIVEN_GUIDE.md](TABLE_DRIVEN_GUIDE.md) |
| レガシー着手の判断フロー | [LEGACY_REFACTORING_APPROACH.md](LEGACY_REFACTORING_APPROACH.md) |
| クロージャによるループ抽象化 | [REFACTORING_CLOSURE.md](REFACTORING_CLOSURE.md) |

---

## 開発者向け: VBA Runner 本体の詳細情報

このリポジトリ（エンジン・拡張・配布物）を開発・改修する人向けです。パッケージを利用するだけの場合は上の「利用者向け」を見てください。

| 目的 | 文書 |
|---|---|
| リポジトリ開発（セットアップ・コマンド） | [CONTRIBUTING.md](../CONTRIBUTING.md) |
| API・制約・Sandbox などの詳細仕様 | [REFERENCE.md](../REFERENCE.md) |
| LSP / VS Code 拡張の実装 | [LSP.md](../LSP.md) |
| エンジン内部設計・仕様メモ | [internals/](internals/) |
| このリポジトリを AI が開発するときの契約 | [CLAUDE.md](../CLAUDE.md) |
