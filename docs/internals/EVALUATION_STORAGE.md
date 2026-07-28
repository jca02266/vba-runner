---
type: concept
title: 評価記録とバグ探索状態の永続化
description: evaluate-vba-runnerの評価履歴、探索キュー、横展開結果、カバレッジ参照を長期管理する方式
tags:
  - evaluation
  - bug-hunting
  - workflow
status: proposed
generated: 2026-07-28
---

# 評価記録とバグ探索状態の永続化

## 目的

`EVAL_LOG.md` に混在している評価履歴、次の探索候補、横展開調査、実Excel照合待ち、
カバレッジメモを分離し、次の4点を同時に満たす。

1. サブエージェントへ渡す情報を絞り、トークンを節約する。
2. 未探索で影響の大きい経路を優先し、1回の評価で得られるバグ発見量を高める。
3. 1つのバグの原因・類似経路・除外経路を次の評価で再利用する。
4. 評価を繰り返しても、未確認候補が減り、確認済み境界が蓄積される状態にする。

## 正本の構成

```text
evaluation/
  evaluations/EV-xxxxx.md   # 1評価1ファイル。YAML + 詳細本文
  campaigns/*.yml           # 探索候補の静的定義
  states/*.yml              # claim、実行中、中断、回収の状態
  coverage-index.yml         # 既存coverage JSONへの参照だけを保持
  schema.yml                # frontmatterと状態の検証規則
  legacy/EVAL_LOG.md        # 移行前ログの保全コピー
```

`EVAL_LOG.md` は正本ではなく、上記のデータから生成するレビュー用サマリーとする。
既存の `coverage-v8/` と `coverage-chunks/` のJSONは形式を変更しない。

### 評価ファイル

frontmatterには状態管理に必要な値だけを置く。再現コード、実行結果、根本原因、
横展開の詳細は本文に記録する。

必須メタデータ:

- `id`, `legacyNumber`, `campaign`, `status`, `priority`, `focus`
- `coverageSnapshot`, `findings`, `tests`, `commit`
- `horizontalAudit.confirmed`, `horizontalAudit.ruledOut`,
  `horizontalAudit.unresolved`

日付、ID、コミットは文字列として扱い、YAMLの暗黙型変換に依存しない。
重複キー、未知の状態、必須項目欠落は検証エラーにする。

### 状態

評価と候補の状態は次の列挙値に限定する。

`queued`、`claimed`、`in-progress`、`verified-no-bug`、`bug-found`、`fixed`、
`blocked`、`abandoned`、`known-limit`、`needs-excel`、`retired`

`queued`（未実施）、`needs-excel`（実機待ち）、`known-limit`（恒久制限）、
`blocked`（外部条件待ち）を文章で代用しない。

### カバレッジ

カバレッジJSONは既存ツールの出力をそのまま使い、`coverage-index.yml` には
パス、コミット、生成日時、SHA-256、生成コマンド、長期保存先だけを記録する。
JSONが存在しない場合は `audit` が警告またはエラーにする。

## CLIと状態遷移

`scripts/eval.mjs` に次の操作を実装する。

```text
eval audit
eval next --limit 1
eval context <candidate-id>
eval claim <candidate-id>
eval release <candidate-id> <claim-token>
eval record <evaluation-file>
eval complete <candidate-id> <evaluation-id> <status> <claim-token>
eval validate
eval render
eval migrate
```

- `audit`: ID重複、参照切れ、stale状態、coverage欠損を検査し、期限切れclaimを回収する。
- `next`: カバレッジ未通過、優先度、未実施状態、既知原因との近接度から候補を1件選ぶ。
- `context`: 候補、関連評価、原因キー、横展開結果、カバレッジだけを出力する。
- `claim`: atomic createと秘密token・所有者・時刻・TTLで二重取得を防ぐ。
- `release`: claim tokenを検証して中断・期限切れの候補を再キューする。
- `record`: 同じIDを再記録しても壊れない冪等更新を行う。
- `complete`: 評価結果を候補へ関連付け、完了候補を次回選定から除外する。
- `validate`: frontmatter、列挙値、必須項目、参照関係を検証する。
- `render`: 決定的な順序で `EVAL_LOG.md` を生成する。
- `migrate`: 旧ログから構造化記録を作る。

評価に失敗して停止した場合も記録を消さず、`abandoned` または `blocked` に遷移させる。
TTL切れの `claimed` / `in-progress` は `release` で回収できる。

## トークン節約と発見効率

通常の評価でサブエージェントへ渡すのは、全履歴ではなく次の最小コンテキストとする。

1. `eval next` が返す候補1件。
2. その候補に紐づく既知バグ、横展開結果、除外経路。
3. 対象ソースと最新カバレッジの未通過箇所。
4. 同じキャンペーンの直近評価だけ。

全履歴が必要なのは移行、監査、重複確認、収束判定のときだけとする。
これにより毎回900行超のMarkdownを読み込むことを避ける。

候補選定では、単なる未実施数ではなく次を優先する。

1. 最新カバレッジで未通過の高影響経路。
2. 直近のバグ原因と同じ共通処理を通る未監査経路。
3. 実Excelとの差分が未確認で、入力を自動生成できる境界。
4. 既存回帰テストでは到達しない組み合わせ。

## 横展開調査の再利用

バグ記録には、原因を抽象化した `causeKey` を付ける。
次回は同じ `causeKey` を持つ未監査のソース経路を自動的に候補化する。

横展開結果は必ず次の3分類で記録する。

- `confirmed`: 同じ原因を確認した経路
- `ruledOut`: 集中テストで問題なしと確認した経路
- `unresolved`: 実Excel確認など、まだ結論を出せない経路

`unresolved` をバグ件数へ加えず、実機照合キューへ送る。

候補に設定した原因仮説は `priorCauseKey`、評価で確定した原因は
`causeKey`（発見原因）として分離する。候補の静的statusと結果ファイルの実効statusは
別に保持し、`context` と `render` では実効statusと関連評価IDを表示する。

## 移行計画

### Phase 1: CLIと試験移行

1. YAMLパーサーと固定スキーマを追加する。
2. `validate`、`audit`、`next`、`claim`、`render` の最小実装を作る。
3. 評価#180〜#188のみを試験移行する。
4. 生成した `EVAL_LOG.md` と旧ログを比較する。
5. 壊れたfrontmatter、重複ID、競合claim、stale回収、coverage欠損のテストを追加する。

### Phase 2: 全履歴移行

1. 評価#100以降を `legacyNumber` を保持したまま移行する。
2. 自由記述は本文または `legacy/notes` に保存する。
3. 旧 `EVAL_LOG.md` は移行前の保全コピーとして残す。
4. 生成ビューの内容をレビューしてから正本を切り替える。

### Phase 3: 評価スキル統合

評価ループは必ず次の順序で実行する。

1. `eval audit`
2. `eval next`
3. `eval claim`
4. 独立評価
5. バグの独立再現
6. 横展開調査
7. 修正・回帰テスト
8. 構造化評価記録の保存
9. `eval validate`
10. `eval render`
11. コミット

構造化記録が保存されていない評価、横展開結果がない修正、回帰テストのない修正は
コミット不可とする。修正なしの場合も `verified-no-bug`、`known-limit`、
`needs-excel` のいずれかを明示して記録する。

## 収束判定

評価ループを終了できるのは、次の条件を満たした場合だけとする。

- `queued` の高優先度候補がない
- 同じ `causeKey` の未監査経路がない
- `needs-excel` の実機照合結果が記録済み、または恒久制限へ分類済み
- 最新カバレッジの高影響未通過箇所に評価候補が登録されている
- `blocked` / `abandoned` が放置されていない

単一のファザー、ミューテーション、カバレッジ実行だけでは収束と判定しない。
