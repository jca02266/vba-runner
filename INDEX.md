# INDEX.md — ドキュメント一覧・評価・再編提案

プロジェクト全体のドキュメントを区分・整理し、評価と再編提案をまとめる。

---

## ドキュメント一覧

### ルート直下

| ファイル | 区分 | 概要 | 行数 |
|---|---|---|---|
| `README.md` | ユーザー向け | プロジェクト概要・クイックスタート・ディレクトリ構造 | 82 |
| `REFERENCE.md` | ユーザー向け | 詳細仕様（型システム・モック・Sandbox・VFS・スコープ・制限事項） | 683 |
| `CLAUDE.md` | AI エージェント向け | Claude Code 向け開発ガイド（コマンド・アーキテクチャ・テストパターン） | 148 |
| `AGENTS.md` | AI エージェント向け | コーディングエージェント向けガイド（英語・エンジン実装担当向け） | 47 |
| `FOR_AI.md` | AI エージェント向け | リファクタリング支援 AI 向けガイド（禁止事項・サイクル手順） | 392 |
| `LSP.md` | 開発者向け | Language Server Protocol 実装仕様・機能一覧 | 318 |
| `TODO.md` | 開発者向け | 次期開発ロードマップ（Phase 2: VS Code 統合） | 1283 |
| `TODO_SPEC.md` | 開発者向け | VBA 言語仕様実装 TODO（MS-VBAL 準拠・ランタイム挙動） | 863 |
| `INDEX.md` | ナビゲーション | 本ファイル：全ドキュメント一覧・評価・再編提案 | — |

### build/ — 配布パッケージ別

| ファイル | 区分 | 概要 |
|---|---|---|
| `build/runner/README.md` | ユーザー向け | npm パッケージ `vba-runner` の使い方・CLI ツール |
| `build/extension/README.md` | ユーザー向け | VS Code 拡張機能のインストール・LSP 機能・設定 |
| `build/playground/README.md` | ユーザー向け | VBA Web Runner（Web UI デモ）の使い方・ローカル起動手順 |

### docs/ — リファクタリング・テスト支援

| ファイル | 区分 | 概要 |
|---|---|---|
| `docs/INDEX.md` | ナビゲーション | docs/ 内ドキュメントの索引・読む順序・用途別ガイド |
| `docs/TESTING_STRATEGY.md` | 基礎 | VBA テスト設計 6 原則（Domain Logic と I/O の分離） |
| `docs/REFACTORING_GUIDE.md` | 基礎 | テスト可能化のためのリファクタリング 4 パターン |
| `docs/TEST_FRAMEWORK_GUIDE.md` | 実践 | JS テストフレームワークで VBA をテスト（VBARunner API リファレンス含む） |
| `docs/MOCK_GUIDE.md` | 実践 | Excel オブジェクトモック実装ガイド（3 段階） |
| `docs/REFACTORING_EXAMPLE.md` | 実例 | TaskScheduler マクロのリファクタリング実例 |
| `docs/INTEGRATION_TEST_EXAMPLE.md` | 実例 | TaskScheduler の統合テスト実装例 |
| `docs/REFACTORING_TESTING_CATALOG.md` | カタログ | リファクタリング・テスト手法の横断参照カタログ（R/T/S/VA コード） |
| `docs/TABLE_DRIVEN_GUIDE.md` | 手法 | テーブル駆動リファクタリング（複雑な分岐を統一テーブルに） |
| `docs/REFACTORING_CLOSURE.md` | 手法 | クロージャによるループ抽象化（JS 版・VBA クラス版対比） |
| `docs/LEGACY_REFACTORING_APPROACH.md` | 手法 | レガシー VBA への対処法（Feathers の考え方を VBA に適用） |
| `docs/FEATHERS_TECHNIQUES_TODO.md` | 手法 | Feathers "Working Effectively with Legacy Code" 手法一覧 |
| `docs/VSCODE_SAMPLE_WORKSPACE.md` | 補足 | `sample/workspace/` パッケージの説明 |

### docs/internals/ — エンジン内部設計

| ファイル | 区分 | 概要 |
|---|---|---|
| `docs/internals/TYPE_SYSTEM_SPEC.md` | 設計仕様 | VBA 型システム仕様（TypeName/VarType/オーバーフロー検査） |
| `docs/internals/NAME_RESOLUTION.md` | 設計仕様 | 名前解決の仕様と実装（MS-VBAL §5.6.10） |
| `docs/internals/DATAFLOW_ANALYSIS_TODO.md` | 計画 | データフロー解析の実装計画 |
| `docs/internals/IDENTIFIER_CLASSIFICATION_TODO.md` | 計画 | 識別子分類の正規化 TODO |

### sample/ — サンプルコード

| ファイル | 区分 | 概要 |
|---|---|---|
| `sample/workspace/README.md` | サンプル | VBA サンプル - Excel データ処理の説明 |
| `sample/excel/README.md` | サンプル | xlsm VBA 抽出・書き戻しツールの説明 |
| `sample/src/refactoring/TaskScheduler.md` | サンプル | TaskScheduler マクロの仕様説明 |
| `sample/src/refactoring/refactoring_quality/core-vs-v5-comparison.md` | 分析 | リファクタリング前後の比較レポート |
| `sample/src/refactoring/refactoring_quality/refactoring-comparison.md` | 分析 | リファクタリング品質比較レポート |

### .claude/ — Claude Code 設定

| ファイル | 区分 | 概要 |
|---|---|---|
| `.claude/commands/implement-vba.md` | AI 向け | VBA 仕様実装・バグ修正の手順コマンド |

---

## ドキュメント評価

論理性・実用性・読みやすさ・独自性・明確性の5軸で評価する。

---

### README.md の評価

#### 論理性（B）

「目的3箇条 → パッケージ表 → クイックスタート → ディレクトリ構成 → ドキュメントリスト」の流れは自然。ただし以下の論理的なギャップがある。

- **クイックスタートの対象が不明確**: step 3 が `npx tsx sample/tests/...`（サンプルテストの実行）、step 4 が `npm test`（仕様テスト一括実行）。「何を確認しているのか」が示されていないため、初見者はなぜこれを実行するのかわからない
- **npm ユーザーへの導線がない**: クイックスタートが `git clone` から始まっており、`npm install vba-runner` だけで利用したいユーザーを想定していない
- **ドキュメントリストに「なぜ読むか」がない**: 末尾のリンク表は並列列挙のみで、目的別の読み分けが示されていない

#### 実用性（B）

- 目的別エントリーポイントがない。「テストを書きたい」「リファクタリングしたい」「LSP を使いたい」それぞれで最初に参照するドキュメントが異なるが、README.md ではすべて同じ順で読む前提になっている
- `docs/` へのリンクが1行（「詳細設計ガイド・実践例」）しかなく、中に12本のドキュメントがあることが伝わらない
- 「リファクタリング対象の VBA ソースを `vba-runner/` と同じ階層に配置し…」という一文はキーユースケースを示しているが、ディレクトリ構成例の後に唐突に現れる

#### 読みやすさ（A-）

- 82行というボリュームは適切で、詳細は REFERENCE.md に委ねる設計は正しい
- ディレクトリ構成の1行コメントは簡潔で読みやすい。ただし `test-libs/` の役割（VBARunner クラスの実装場所・CLI ツール群の本体）がコメントから読み取れない
- 「VBA ソースとテストを並べて配置する場合の例」のディレクトリ図はユースケースをイメージさせる点で実用的

#### 独自性（C）— 最大の改善余地

- 「Excel 不要」という差別化は冒頭にあるが、競合（Rubberduck・VBA-Unit 等）との違いが示されていない
- **TODO.md の「VBA Runner なしの場合との比較」表**は「仮説から検証済みへ」という強力なバリュープロポジションを持つが、README.md では「AIの支援を受けながら」の1文で終わっている。このバリューが README.md に反映されていないことが独自性評価を下げる最大の要因
- なぜ VBA を TypeScript でテストするのか、という根本的な問いへの回答が弱い

#### 明確性（B）

- **表記揺れ（修正済み）**: 「提供パッケージ」テーブルで `vba-runner`（npm ライブラリ）・`VBA Runner`（VS Code 拡張機能）・`VBA Web Runner`（Web UI）と表記スタイルが混在していたが修正対象
- **ドキュメントテーブルの `FOR_AI.md`**: 「AIによるリファクタリング支援のガイド」と書かれているが、人間ユーザーが自分事として読む理由が不明確
- クイックスタートの step 3・4 の目的（動作確認なのか仕様確認なのか）が曖昧

---

### ドキュメント構成全体の評価

#### 論理性（B+）

プロジェクトの2軸（「VBA実行エンジン実装」と「リファクタリング支援」）に対してドキュメントが分かれており、方向性は正しい。

```
VBA実行エンジン軸: CLAUDE.md → REFERENCE.md → LSP.md → TODO.md / TODO_SPEC.md
リファクタリング支援軸: FOR_AI.md → docs/ (TESTING_STRATEGY → REFACTORING_GUIDE → ...)
```

- **問題**: 2軸の入口が README.md で統一されていないため、用途に応じた読み始めが案内されない
- **問題**: `docs/INDEX.md` は `docs/` 内の学習パスを丁寧に示しているが、`docs/internals/` は対象外であり、ルート直下の `AGENTS.md` / `FOR_AI.md` との接続も示されていない

#### 実用性（B+）

- `docs/` の学習体系（TESTING_STRATEGY → REFACTORING_GUIDE → TEST_FRAMEWORK_GUIDE → MOCK_GUIDE）は整然としており、シナリオ別の読み方まで `docs/INDEX.md` が案内している点は高品質
- **問題**: eval/run の説明が `build/runner/README.md`（基本的な使い方）と `REFERENCE.md`（詳細な使い分け）の2か所に分散しており、どちらをどの深度で参照すればよいか不明確
- **問題**: `AGENTS.md` の Sandbox 方針が `README.md` を参照しているが、実際の情報は `REFERENCE.md` に移動済みのため、参照が機能していない（→ 修正対象）

#### 読みやすさ（B）

- ルート直下に AI 向けファイルが3本（`CLAUDE.md`・`AGENTS.md`・`FOR_AI.md`）あり、人間ユーザーが何を読めばよいかを迷う
- `TODO.md`（1283行）・`TODO_SPEC.md`（863行）がルート直下に置かれているが、「開発者が更新するリビングドキュメント」なのか「仕様書」なのかが外から判断しにくい

各ドキュメントの対象読者明示度のばらつき：

| ファイル | 対象読者の明示 |
|---|---|
| `FOR_AI.md` | 明示あり（「リファクタリング支援 AI 向け」） |
| `CLAUDE.md` | ファイル名で暗示（Claude Code 向け） |
| `AGENTS.md` | 英語冒頭で暗示（エンジン実装担当エージェント向け） |
| `LSP.md` | 明示なし（VS Code 拡張開発者が主な対象） |
| `TODO.md` | 明示なし（プロジェクトオーナー向けのロードマップ） |

#### 独自性（A-）

- `docs/` のリファクタリング支援ドキュメント群は質が高く、競合ツールには存在しないレベルの実践ガイドになっている
- `FOR_AI.md` のリファクタリングサイクル（Phase 1-7）と `vba-analyzer` を連動させた手順は独自価値が高い
- **問題**: これらの強みが README.md のエントリーポイントで伝わっていない。プロジェクトの「ユニークさ」がドキュメント入口と本文の間で断絶している

#### 明確性（B-）— 構成面での最大課題

- AI向けファイル3本が「誰が何のために読むか」という軸で整理されておらず、初見者は用途を判断しにくい
- `docs/INDEX.md`（docs/ 専用索引）と `INDEX.md`（本ファイル：全体索引）の2層構造になったが、README.md から `INDEX.md` へのリンクがない
- `build/playground/README.md` のタイトルが「VBA Runner Playground」と旧称のままだった（→ 修正対象）

---

## 改善点リスト

評価に基づく具体的な改善点を優先度順に列挙する。

### 高優先（README.md の独自性・実用性の向上）

| # | 対象 | 改善内容 |
|---|---|---|
| 1 | `README.md` 目的セクション | `TODO.md` の「VBA Runner なしの場合との比較」表を要約して追加。「AI提案は仮説→VBA Runnerで即検証」というバリューを明文化 |
| 2 | `README.md` | 用途別エントリーポイントを追加。「テストを書きたい」「リファクタリングしたい」「LSP を使いたい」の3パターンで最初に読むドキュメントへ誘導 |
| 3 | `README.md` | ドキュメントテーブルに「対象読者 / 用途」列を追加し、リンク集から脱却 |
| 4 | `README.md` | `INDEX.md`（全ドキュメント索引）へのリンクを追加 |

### 高優先（構成の明確性・整合性の修正）

| # | 対象 | 改善内容 |
|---|---|---|
| 5 | `AGENTS.md` | Sandbox 方針の参照先を `README.md` → `REFERENCE.md` に修正（陳腐化した参照の解消） |
| 6 | `README.md` 提供パッケージ表 | 「Playground」→「VBA Web Runner」に名称統一 |
| 7 | `build/playground/README.md` | タイトルを「VBA Web Runner — Web UI デモ」に修正 |

### 中優先（構成の整理・統合）

| # | 対象 | 改善内容 |
|---|---|---|
| 8 | `AGENTS.md` | `CLAUDE.md` に統合して廃止。エンジン実装向けのスタブ方針・Sandbox 方針は CLAUDE.md の適切なセクションへ移動 |
| 9 | `build/runner/README.md` と `REFERENCE.md` | eval/run 説明の重複を整理。build/runner/README.md は概要・典型例のみ、詳細比較は REFERENCE.md に一元化 |
| 10 | `docs/INDEX.md` | `docs/internals/` のドキュメントを追加（現状は docs/ 直下のみをカバー） |

---

## 全体評価

| 観点 | README.md | 構成全体 | 主な根拠 |
|---|---|---|---|
| 論理性 | B | B+ | 大きな流れは正しいが、2軸の入口分岐がなく目的別導線が弱い |
| 実用性 | B | B+ | docs/ の学習体系は高品質。npm ユーザー向け・目的別の導線が弱点 |
| 読みやすさ | A- | B | README.md 単体は適切なボリューム。AI向けファイル3本の混在が全体を読みにくくしている |
| 独自性 | C | A- | TODO.md に強力なバリュープロポジションがあるが README.md に未反映 |
| 明確性 | B | B- | パッケージ名の表記揺れ・対象読者の不統一・docs/internals/ のインデックス漏れが課題 |

**強み**:
- `docs/` のリファクタリング・テスト支援体系は内容・構成ともに高品質
- `FOR_AI.md` の Phase 1-7 サイクルと `vba-analyzer` 連動は独自価値が高い
- `build/` 配下の配布物別 README 分離は整理されており適切

**最大の課題**:
1. **独自性の表面化不足** — プロジェクト最大の強み「AI + VBA実行エンジン + リファクタリング支援」の組み合わせが README.md の入口で伝わっていない
2. **AI向けドキュメントの散在** — `CLAUDE.md` / `AGENTS.md` / `FOR_AI.md` の3本が混在し、人間ユーザーが迷う入口を作っている

---

## 再編提案

評価を踏まえた具体的なアクション計画。

### フェーズ1: 即実行（壊れた参照・名称の修正）

| アクション | 対象 | 内容 |
|---|---|---|
| 修正 | `AGENTS.md` | Sandbox 方針の参照先を `README.md` → `REFERENCE.md` に修正 |
| 修正 | `README.md` | 「Playground」→「VBA Web Runner」に名称統一 |
| 修正 | `build/playground/README.md` | タイトルを「VBA Web Runner — Web UI デモ」に修正 |
| 追加 | `README.md` | ドキュメントテーブルに `INDEX.md` へのリンクを追加 |

### フェーズ2: README.md の強化（独自性・実用性の向上）

README.md はプロジェクトへの最初の入口であり、独自性（C評価）の改善が最も効果が高い。

| アクション | 内容 |
|---|---|
| 目的セクション強化 | 「VBA Runner なしの場合との比較」を要約。「AI提案→即検証」のループが高速化するという核心価値を1-2文で表現 |
| 用途別エントリー追加 | テスト・リファクタリング・LSP の3ルートで最初のドキュメントへ誘導する小セクションを追加 |
| ドキュメントテーブル改善 | `FOR_AI.md` の説明を人間ユーザー目線に変更（「VBA リファクタリングを AI に依頼する場合の操作手順」等）。用途列を追加 |

### フェーズ3: AI向けドキュメントの整理（明確性の向上）

3本の AI 向けファイルを2本に整理する。

```
現状:
  CLAUDE.md  ← Claude Code 向け（日本語・プロジェクト全般）
  AGENTS.md  ← エンジン実装担当エージェント向け（英語・Sandbox/スタブ方針）
  FOR_AI.md  ← リファクタリング支援 AI 向け（日本語・リファクタリングサイクル）

提案後:
  CLAUDE.md  ← Claude Code 向け（AGENTS.md の内容を統合）
  FOR_AI.md  ← リファクタリング専用（変更なし）
  AGENTS.md  ← 廃止（内容を CLAUDE.md へ吸収）
```

AGENTS.md の吸収先: `CLAUDE.md` の「ファイル入出力の実装規則」セクション付近にスタブ方針・Sandbox方針を追記する。

### フェーズ4: インデックスの整合性向上（読みやすさの向上）

| アクション | 対象 | 内容 |
|---|---|---|
| 追加 | `docs/INDEX.md` | `docs/internals/` の4ファイルをインデックスに追加 |
