---
type: concept
title: 評価記録とバグ探索状態の永続化
description: evaluate-vba-runnerの評価履歴、探索キュー、横展開結果、カバレッジ参照を長期管理する現行方式
tags:
  - evaluation
  - bug-hunting
  - workflow
status: active
generated: 2026-07-29
---

# 評価記録とバグ探索状態の永続化

## この文書の範囲

この文書は、`evaluate-vba-runner` の評価履歴と候補キューを、現在どのように
保存・選定・完了管理するかを説明する。VBAの仕様準拠TODOは
[`TODO_SPEC.md`](../../TODO_SPEC.md)、仕様準拠外の開発課題は
[`TODO.md`](../../TODO.md)で管理する。

## 当初の問題

当初は、評価結果を1つの `EVAL_LOG.md` に自由記述で追記していた。この方式では、
次の情報が同じMarkdownに混在していた。

- 評価の実施結果とバグ修正履歴
- 次に試す候補、実施中の候補、実機照合待ち
- バグの原因と、同じ原因を調査した経路
- カバレッジの未通過箇所と、次の重点領域

その結果、次の問題が発生していた。

1. 毎回全文を読み込む必要があり、評価プロンプトのトークンを消費した。
2. 「未実施」「実機待ち」「恒久的制限」が文章表現になり、状態を機械的に判定できなかった。
3. 同じ候補を複数エージェントが取得したり、完了済み候補を再評価したりしやすかった。
4. バグの横展開結果が次の評価に再利用されず、同じソース経路を繰り返し調査した。
5. coverageのメモと評価結果の対応が曖昧で、低カバレッジ領域を優先できなかった。

## 見直しで行った改善

自由記述ログを置き換えるにあたり、1評価1ファイルのYAML frontmatter付きMarkdownを
正本とした。カバレッジJSON自体は既存形式を維持し、参照情報だけを別ファイルに置いた。

### 記録と状態を分離

- `evaluation/evaluations/EV-xxxxx.md` に評価内容を保存する。
- `evaluation/campaigns/*.yml` に探索候補を定義する。
- `evaluation/states/*.yml` にclaimと完了結果を保存する。
- `evaluation/coverage-index.yml` にcoverageスナップショットを参照する。
- `evaluation/schema.yml` で必須項目と状態を検証する。

### 状態遷移をCLI化

`scripts/eval.mjs` に監査、候補選定、claim、解放、記録、完了、検証、描画を実装した。
claimは秘密token・所有者・TTL付きのatomic createで二重取得を防ぐ。完了時は、評価の
`candidateId` と候補ID、campaign、状態、評価IDを照合するため、同一campaign内の
別評価を誤って完了登録できない。

resultについても、次を `validate` / `audit` で検証する。

- resultファイル名と内部 `candidateId` の一致
- candidate、evaluation、campaignの存在と対応
- resultの状態と評価の状態の一致

### 原因と横展開を再利用

候補の仮説は `priorCauseKey`、評価で確認した原因は `causeKey` として分ける。
横展開調査は `confirmed`、`ruledOut`、`unresolved` の3分類で記録する。
`unresolved` はバグ件数に含めず、実Excel照合などの待ち状態として扱う。

### coverageで候補を優先

`next` は優先度だけでなく、最新coverageスナップショットの未通過対象と候補の
`coverageTargets` を照合する。候補選定は、低カバレッジの高影響経路、直近の原因と
共通化された経路、実Excel差分、未到達の組み合わせの順に考える。

### 履歴を移行

- 評価#100〜#190を移行し、その後EV-00191を追加した（現在92件）。
- 旧ログは `evaluation/legacy/EVAL_LOG.md` に保全した。
- ルートの `EVAL_LOG.md` は正本ではなく、構造化記録から生成するレビュー用ビューにした。
- 移行した履歴にも安定した `candidateId` を付与し、新規記録と同じスキーマで検証できるようにした。

## 現在の正本構成

```text
evaluation/
  evaluations/EV-xxxxx.md   # 1評価1ファイル。YAML + 詳細本文
  findings/BUG-xxxxx.md      # 評価で確認したバグ
  campaigns/*.yml            # 探索候補と優先度・coverage対象
  states/*.claim.yml         # 実行中claim（TTL付き）
  states/*.result.yml        # 完了した候補と評価ID・状態
  coverage-index.yml         # 既存coverage JSONへの参照
  schema.yml                 # frontmatterと状態の検証規則
  legacy/EVAL_LOG.md         # 移行前ログの保全コピー
```

`EVAL_LOG.md` は手編集しない。必要な場合は次を実行して再生成する。

```bash
npm run eval -- render EVAL_LOG.md
```

coverage JSONは `coverage-v8/` や `coverage-chunks/` にある既存出力をそのまま使い、
`coverage-index.yml` にはスナップショットID、生成日時、コミット、対象レポート、
未通過ファイルを記録する。

## 評価記録の要件

frontmatterには機械的に扱う情報を置き、再現コード、実行結果、原因、横展開の詳細は
本文に置く。全評価で次の項目を記録する。

- `id`, `candidateId`, `campaign`, `status`, `priority`, `focus`
- `coverageSnapshot`, `findings`, `tests`, `commit`
- `horizontalAudit.confirmed`, `horizontalAudit.ruledOut`,
  `horizontalAudit.unresolved`

バグを修正した評価は、原因キー、回帰テスト、修正コミットを必ず記録する。
バグがない場合も `verified-no-bug`、実機照合待ちは `needs-excel`、恒久的制限は
`known-limit`、中断は `abandoned` または `blocked` として記録する。

使用可能な状態は次のとおりである。

`queued`、`claimed`、`in-progress`、`verified-no-bug`、`bug-found`、`fixed`、
`blocked`、`abandoned`、`known-limit`、`needs-excel`、`retired`

## 現行CLIと役割

```text
npm run eval -- audit
npm run eval -- next --limit 1
npm run eval -- context <candidate-id>
npm run eval -- claim <candidate-id>
npm run eval -- release <candidate-id> <claim-token>
npm run eval -- record <evaluation-file>
npm run eval -- complete <candidate-id> <evaluation-id> <status> <claim-token>
npm run eval -- validate
npm run eval -- render EVAL_LOG.md
```

- `audit`: 記録、campaign、result、coverage参照を検証し、期限切れclaimを回収する。
- `next`: 実施可能な候補から、優先度とcoverage一致を考慮して選ぶ。
- `context`: 候補、関連評価、原因キー、横展開結果、coverageだけを出力する。
- `claim`: 候補を予約し、tokenを発行する。
- `release`: tokenを検証してclaimを解放する。
- `record`: 評価記録を冪等に保存する。同じIDの異なる内容は拒否する。
- `complete`: 候補と評価の対応、状態、修正コミットとテストを検証して完了結果を保存する。
- `validate`: frontmatter、必須項目、参照関係、result整合性を検証する。
- `render`: 構造化記録から決定的なMarkdownビューを生成する。

## 現行の評価手順

### 1. 候補を選ぶ

```bash
npm run eval -- audit
npm run eval -- next --limit 1
npm run eval -- context <candidate-id>
npm run eval -- claim <candidate-id>
```

サブエージェントには全履歴を渡さず、`next` と `context` の出力、対象ソース、
最新coverage、同じcampaignの直近評価だけを渡す。評価対象のVBAとドライバーは
リポジトリ外（通常は `/tmp/vba-runner-eval/`）に置く。

### 2. 独立評価と再現

独立評価エージェントに、公式READMEを読んだ新規利用者としてサンプルVBAを実行させる。
バグ報告はそのまま採用せず、最小再現を自分で実行して確認する。

バグが再現した場合は、修正前に別エージェントで同種のdispatch・評価経路を横展開する。
確認済み、除外済み、実Excelなどで未確定の経路を評価本文に記録する。

### 3. 修正または結果記録

- バグあり: 原因を特定し、最小修正、回帰テスト、評価記録を作る。
- バグなし: `verified-no-bug` として、試した境界と除外経路を記録する。
- 実機待ち: `needs-excel` として、照合内容を `unresolved` に記録する。
- 恒久制限: `known-limit` として、現状の制限と解決時の効果を記録する。

修正した場合は、評価記録の保存後に次を実行する。

```bash
npm run eval -- validate
npm run eval -- render EVAL_LOG.md
git add <implementation> <tests> evaluation EVAL_LOG.md
git commit
```

評価が中断した場合はファイルを削除せず、claimを解放するか、`abandoned` / `blocked`
として次回の `audit` で回収できるようにする。

### 4. 次の候補へ進む

完了後に `complete` で候補と評価を関連付ける。結果ファイルが作成されると、
`next` はその候補を自動的に除外する。次の評価では、直近の `causeKey` と横展開結果を
再利用し、同じテストを理由なく繰り返さない。

## トークン節約と収束

全履歴が必要なのは移行、重複確認、監査、収束判定だけである。通常の評価では候補の
compact contextだけを渡すため、旧ログ全文を毎回読み込まない。

評価ループを終了できるのは、次の条件をすべて満たした場合である。

- 高優先度の `queued` 候補がない
- 直近の `causeKey` に属する未監査経路がない
- `needs-excel` の照合結果が記録済み、または恒久制限に分類済み
- 最新coverageの高影響未通過箇所に候補が登録されている
- `blocked` / `abandoned` / stale claim が放置されていない

単一のファザー、ミューテーション、coverage実行だけでは収束と判定しない。
