# INDEX.md — ドキュメント一覧

プロジェクト全体のドキュメントを区分・整理し、評価と再編提案をまとめる。

---

## ドキュメント一覧

### ルート直下

| ファイル | 区分 | 概要 | 規模 |
|---|---|---|---|
| [README.md](README.md) | ユーザー向け | プロジェクト概要・クイックスタート・ディレクトリ構造 | 小 |
| [REFERENCE.md](REFERENCE.md) | ユーザー向け | 詳細仕様（型システム・モック・Sandbox・VFS・スコープ・制限事項） | 大 |
| [CLAUDE.md](CLAUDE.md) | AI エージェント向け | Claude Code 向け開発ガイド（コマンド・アーキテクチャ・テストパターン） | 小 |
| [FOR_AI.md](FOR_AI.md) | AI エージェント向け | リファクタリング支援 AI 向けガイド（禁止事項・サイクル手順） | 中 |
| [LSP.md](LSP.md) | 開発者向け | Language Server Protocol 実装仕様・機能一覧 | 中 |
| [TODO.md](TODO.md) | 開発者向け | 次期開発ロードマップ（Phase 2: VS Code 統合） | 大 |
| [TODO_SPEC.md](TODO_SPEC.md) | 開発者向け | VBA 言語仕様実装 TODO（MS-VBAL 準拠・ランタイム挙動） | 大 |
| [INDEX.md](INDEX.md) | ナビゲーション | 本ファイル：全ドキュメント一覧 | — |

### build/ — 配布パッケージ別

| ファイル | 区分 | 概要 | 規模 |
|---|---|---|---|
| [build/runner/README.md](build/runner/README.md) | ユーザー向け | npm パッケージ `vba-runner` の使い方・CLI ツール | 小 |
| [build/extension/README.md](build/extension/README.md) | ユーザー向け | VS Code 拡張機能のインストール・LSP 機能・設定 | 小 |
| [build/playground/README.md](build/playground/README.md) | ユーザー向け | VBA Web Runner（Web UI デモ）の使い方・ローカル起動手順 | 小 |

### docs/ — リファクタリング・テスト支援

| ファイル | 区分 | 概要 | 規模 |
|---|---|---|---|
| [docs/INDEX.md](docs/INDEX.md) | ナビゲーション | docs/ 内ドキュメントの索引・読む順序・用途別ガイド | 中 |
| [docs/TUTORIAL.md](docs/TUTORIAL.md) | 入門 | vba-extractor 抽出 → VS Code 確認 → AI リファクタリング/機能追加 → 書き戻しの一連の流れを実例で体験するチュートリアル | 中 |
| [docs/TUTORIAL_AI.md](docs/TUTORIAL_AI.md) | 入門 | 同じ流れを、コマンドの詳細に触れず AI への指示文だけで完結させるチュートリアル | 小 |
| [docs/TESTING_STRATEGY.md](docs/TESTING_STRATEGY.md) | 基礎 | VBA テスト設計 6 原則（Domain Logic と I/O の分離） | 大 |
| [docs/REFACTORING_GUIDE.md](docs/REFACTORING_GUIDE.md) | 基礎 | テスト可能化のためのリファクタリング 4 パターン | 大 |
| [docs/TEST_FRAMEWORK_GUIDE.md](docs/TEST_FRAMEWORK_GUIDE.md) | 実践 | JS テストフレームワークで VBA をテスト（VBARunner API リファレンス含む） | 大 |
| [docs/MOCK_GUIDE.md](docs/MOCK_GUIDE.md) | 実践 | Excel オブジェクトモック実装ガイド（3 段階） | 大 |
| [docs/REFACTORING_EXAMPLE.md](docs/REFACTORING_EXAMPLE.md) | 実例 | TaskScheduler マクロのリファクタリング実例 | 大 |
| [docs/INTEGRATION_TEST_EXAMPLE.md](docs/INTEGRATION_TEST_EXAMPLE.md) | 実例 | TaskScheduler の統合テスト実装例 | 中 |
| [docs/REFACTORING_TESTING_CATALOG.md](docs/REFACTORING_TESTING_CATALOG.md) | カタログ | リファクタリング・テスト手法の横断参照カタログ（R/T/S/VA コード） | 大 |
| [docs/TABLE_DRIVEN_GUIDE.md](docs/TABLE_DRIVEN_GUIDE.md) | 手法 | テーブル駆動リファクタリング（複雑な分岐を統一テーブルに） | 中 |
| [docs/REFACTORING_CLOSURE.md](docs/REFACTORING_CLOSURE.md) | 手法 | クロージャによるループ抽象化（JS 版・VBA クラス版対比） | 中 |
| [docs/LEGACY_REFACTORING_APPROACH.md](docs/LEGACY_REFACTORING_APPROACH.md) | 手法 | レガシー VBA への対処法（Feathers の考え方を VBA に適用） | 中 |
| [docs/FEATHERS_TECHNIQUES_TODO.md](docs/FEATHERS_TECHNIQUES_TODO.md) | 手法 | Feathers "Working Effectively with Legacy Code" 手法一覧 | 小 |
| [docs/VSCODE_SAMPLE_WORKSPACE.md](docs/VSCODE_SAMPLE_WORKSPACE.md) | 補足 | `sample/workspace/` パッケージの説明 | 小 |

### docs/internals/ — エンジン内部設計

| ファイル | 区分 | 概要 | 規模 |
|---|---|---|---|
| [docs/internals/TYPE_SYSTEM_SPEC.md](docs/internals/TYPE_SYSTEM_SPEC.md) | 設計仕様 | VBA 型システム仕様（TypeName/VarType/オーバーフロー検査） | 小 |
| [docs/internals/NAME_RESOLUTION.md](docs/internals/NAME_RESOLUTION.md) | 設計仕様 | 名前解決の仕様と実装（MS-VBAL §5.6.10） | 中 |
| [docs/internals/DATAFLOW_ANALYSIS_TODO.md](docs/internals/DATAFLOW_ANALYSIS_TODO.md) | 計画 | データフロー解析の実装計画 | 中 |
| [docs/internals/IDENTIFIER_CLASSIFICATION_TODO.md](docs/internals/IDENTIFIER_CLASSIFICATION_TODO.md) | 計画 | 識別子分類の正規化 TODO | 小 |

### sample/ — サンプルコード

| ファイル | 区分 | 概要 | 規模 |
|---|---|---|---|
| [sample/workspace/README.md](sample/workspace/README.md) | サンプル | VBA サンプル - Excel データ処理の説明 | 中 |
| [sample/excel/README.md](sample/excel/README.md) | サンプル | xlsm VBA 抽出・書き戻しツールの説明 | 中 |
| [sample/src/refactoring/TaskScheduler.md](sample/src/refactoring/TaskScheduler.md) | サンプル | TaskScheduler マクロの仕様説明 | 小 |
| [sample/src/refactoring/refactoring_quality/core-vs-v5-comparison.md](sample/src/refactoring/refactoring_quality/core-vs-v5-comparison.md) | 分析 | リファクタリング前後の比較レポート | 中 |
| [sample/src/refactoring/refactoring_quality/refactoring-comparison.md](sample/src/refactoring/refactoring_quality/refactoring-comparison.md) | 分析 | リファクタリング品質比較レポート | 小 |

### .claude/ — Claude Code 設定

| ファイル | 区分 | 概要 | 規模 |
|---|---|---|---|
| [.claude/commands/implement-vba.md](.claude/commands/implement-vba.md) | AI 向け | VBA 仕様実装・バグ修正の手順コマンド | 中 |
