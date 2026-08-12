# ドキュメント台帳（メンテナ向け）

> 利用者の入口は [README.md](./README.md) → [docs/README.md](./docs/README.md) です。
> 開発者の入口は [CONTRIBUTING.md](./CONTRIBUTING.md) です。
>
> このファイルは README に載せない文書も含めた **メンテナ向け台帳** です。
> 入口・地図としては使いません。

---

## ルート直下

| ファイル | 対象 | 概要 |
|---|---|---|
| [README.md](./README.md) / [README.ja.md](./README.ja.md) | 全員 | プロジェクト入口（動機・目的選択） |
| [CONTRIBUTING.md](./CONTRIBUTING.md) | 開発者 | セットアップ・コマンド・次に読む資料 |
| [REFERENCE.md](./docs/implementation/REFERENCE.md) | 利用者・開発者 | 詳細仕様（型・モック・Sandbox・VFS など） |
| [FOR_AI.md](./FOR_AI.md) | 利用者（AI） | 利用者の VBA を AI が扱う実行契約 |
| [CLAUDE.md](./CLAUDE.md) | 開発者（AI） | このリポジトリを AI が開発する実行契約 |
| [LSP.md](./docs/implementation/LSP.md) | 開発者 | LSP 実装仕様 |
| [TODO.md](./docs/todo/TODO.md) | 開発者 | 次期開発ロードマップ |
| [TODO_SPEC.md](./docs/todo/TODO_SPEC.md) | 開発者 | VBA 言語仕様実装 TODO |
| [VBA_ENGINE_BUG_HISTORY.md](./docs/bug-records/VBA_ENGINE_BUG_HISTORY.md) | 開発者 | 過去のVBAエンジン仕様バグ履歴 |
| [INDEX.md](./INDEX.md) | メンテナ | 本台帳 |

## build/ — 配布パッケージ正本

| ファイル | 対象 | 概要 |
|---|---|---|
| [build/runner/README.md](./build/runner/README.md) | 利用者 | npm パッケージ `vba-runner` |
| [build/extractor/README.md](./build/extractor/README.md) | 利用者 | CLI `vba-extractor` |
| [build/extension/README.md](./build/extension/README.md) | 利用者 | VS Code 拡張機能 |
| [build/playground/README.md](./build/playground/README.md) | 利用者 | Web UI デモ |

## docs/ — 利用者ガイドと詳細索引

| ファイル | 対象 | 概要 |
|---|---|---|
| [docs/README.md](./docs/README.md) | 全員 | 目的別の共通ハブ（正本の地図） |
| [docs/legacy-index.md](./docs/legacy-index.md) | 利用者 | 読み順・問題別検索（任意の詳細索引） |
| [TUTORIAL.md](./docs/guides/TUTORIAL.md) | 利用者 | 抽出 → 編集 → 書き戻しの一連体験 |
| [TUTORIAL_AI.md](./docs/guides/TUTORIAL_AI.md) | 利用者 | AI 指示だけで同じ流れを進める |
| [TESTING_STRATEGY.md](./docs/testing/TESTING_STRATEGY.md) | 利用者 | テスト設計 6 原則 |
| [REFACTORING_GUIDE.md](./docs/refactoring/REFACTORING_GUIDE.md) | 利用者 | テスト可能化の 4 パターン |
| [TEST_FRAMEWORK_GUIDE.md](./docs/testing/TEST_FRAMEWORK_GUIDE.md) | 利用者 | JS から VBA をテスト（API 含む） |
| [MOCK_GUIDE.md](./docs/guides/MOCK_GUIDE.md) | 利用者 | Excel モック実装 |
| [REFACTORING_EXAMPLE.md](./docs/refactoring/REFACTORING_EXAMPLE.md) | 利用者 | TaskScheduler 実例 |
| [INTEGRATION_TEST_EXAMPLE.md](./docs/testing/INTEGRATION_TEST_EXAMPLE.md) | 利用者 | 統合テスト実例 |
| [REFACTORING_TESTING_CATALOG.md](./docs/testing/REFACTORING_TESTING_CATALOG.md) | 利用者 | 手法カタログ（R/T/S/VA） |
| [TABLE_DRIVEN_GUIDE.md](./docs/refactoring/TABLE_DRIVEN_GUIDE.md) | 利用者 | テーブル駆動リファクタリング |
| [REFACTORING_CLOSURE.md](./docs/refactoring/REFACTORING_CLOSURE.md) | 利用者 | クロージャによるループ抽象化 |
| [LEGACY_REFACTORING_APPROACH.md](./docs/refactoring/LEGACY_REFACTORING_APPROACH.md) | 利用者 | レガシー着手の判断フロー |
| [FEATHERS_TECHNIQUES_TODO.md](./docs/refactoring/FEATHERS_TECHNIQUES_TODO.md) | 利用者 | Feathers 手法一覧 |
| [VSCODE_SAMPLE_WORKSPACE.md](./docs/guides/VSCODE_SAMPLE_WORKSPACE.md) | 利用者 | `sample/workspace/` の説明 |

## docs/internals/ — エンジン内部設計

| ファイル | 対象 | 概要 |
|---|---|---|
| [docs/internals/BNF.md](./docs/internals/BNF.md) | 開発者 | MS-VBAL BNF |
| [docs/internals/TYPE_SYSTEM_SPEC.md](./docs/internals/TYPE_SYSTEM_SPEC.md) | 開発者 | 型システム |
| [docs/internals/LITERALS_AND_OPERATORS.md](./docs/internals/LITERALS_AND_OPERATORS.md) | 開発者 | リテラル・演算子 |
| [docs/internals/NAME_RESOLUTION.md](./docs/internals/NAME_RESOLUTION.md) | 開発者 | 名前解決 |
| [docs/internals/EVALUATION.md](./docs/internals/EVALUATION.md) | 開発者 | 評価パイプライン |
| [docs/internals/FUNCTION_CALL.md](./docs/internals/FUNCTION_CALL.md) | 開発者 | 関数コール・引数 |
| [docs/internals/AMBIGUOUS_SYNTAX.md](./docs/internals/AMBIGUOUS_SYNTAX.md) | 開発者 | 構文的曖昧性 |
| [docs/internals/VBA_EXPORT_IMPORT.md](./docs/internals/VBA_EXPORT_IMPORT.md) | 開発者 | export/import 技術メモ |
| [docs/internals/DECIMAL_IMPLEMENTATION.md](./docs/internals/DECIMAL_IMPLEMENTATION.md) | 開発者 | Decimal 実装メモ |
| [docs/internals/DATAFLOW_ANALYSIS_TODO.md](./docs/internals/DATAFLOW_ANALYSIS_TODO.md) | 開発者 | データフロー解析計画 |
| [docs/internals/IDENTIFIER_CLASSIFICATION.md](./docs/internals/IDENTIFIER_CLASSIFICATION.md) | 開発者 | VBA識別子分類の実装説明 |
| [docs/bug-records/audit-func-log.md](./docs/bug-records/audit-func-log.md) | 開発者 | 組み込み関数の仕様監査履歴 |

## sample/ / .claude/

| ファイル | 対象 | 概要 |
|---|---|---|
| [sample/workspace/README.md](./sample/workspace/README.md) | 利用者 | サンプルワークスペース |
| [sample/excel/README.md](./sample/excel/README.md) | 利用者 | xlsm 抽出・書き戻しサンプル |
| [sample/src/refactoring/TaskScheduler.md](./sample/src/refactoring/TaskScheduler.md) | 利用者 | TaskScheduler 仕様 |
| [sample/src/refactoring/refactoring_quality/](./sample/src/refactoring/refactoring_quality) | メンテナ | リファクタリング比較レポート |
| [.claude/commands/implement-vba.md](.claude/commands/implement-vba.md) | 開発者（AI） | VBA 仕様実装・バグ修正手順 |
