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

候補の仮説は `priorCauseKey` として保存する。バグを確認した場合は原因を一つの
キーに押し込めず、次の二層に分けて保存する。

- `directCauseKey`: 観測された不具合を直接生じさせた実装上の機構・経路。
- `causeKey`: 直接原因を生んだ設計上の真因。横展開のグルーピングにも使う。

`causeKey` は一意なIDではなく再利用可能な分類キーである。評価・バグは
`EV-xxxxx`、`BUG-xxxxx`、真因分析は`RC-xxxxx`、真因対処は`ROOT-xxxxx`で一意にする。
原因分析の対処状況は`directFixStatus` と `rootFixStatus` に分ける。真因対処が完了した
場合は`rootFixTaskId` と `rootFixCommit` を記録する。`rootFixCandidateId`は旧方式との
互換用であり、新しい真因対処では使用しない。

評価記録には`rootCauseProcedureVersion`を付ける。`0`は手順未整備時代の旧記録、
`1`は[`ROOT_CAUSE_ANALYSIS.md`](ROOT_CAUSE_ANALYSIS.md) v1に従った記録である。
番号がない記録は方式不明として確定根拠に使わない。

原因キーだけでは分析の中身が抜け落ちるため、新方式の評価では
`rootCauseAnalysis`を必須の構造化項目とする。次の項目をすべて記録する。

- `status`: `confirmed`、`hypothesis`、`ruled-out`、`not-applicable`
- `directCause`: 直接原因の機構と発生経路
- `designCause`: 直接原因を許した設計上の真因
- `confirmed`: 真因分析で確認した経路・証拠
- `ruledOut`: 同じ原因ではないと除外した経路
- `unresolved`: まだ確認できない経路・境界
- `decision`: 共通化、局所修正、制限、TODO化の判断と次の対処

新方式の`RC-xxxxx`と`ROOT-xxxxx`にも`procedureVersion: 1`を記録し、起点評価の
`rootCauseProcedureVersion`と一致させる。

`horizontalAudit`が挙動の横展開を記録するのに対し、`rootCauseAnalysis`はなぜその
挙動になったかと対処方針を記録する。`scripts/eval.mjs validate`が必須項目と状態値を
検証するため、真因分析を本文だけに書いて完了扱いにしてはならない。EV-00384以前の
移行済み履歴は旧形式として保持し、新規記録はこの項目を省略しない。

横展開調査は `confirmed`、`ruledOut`、`unresolved` の3分類で記録する。
`unresolved` はバグ件数に含めず、実Excel照合などの待ち状態として扱う。

### 真因仮説と対処タスクを分離

最初の真因分析は確定情報ではない。`evaluation/root-causes/RC-xxxxx.md`へ
`hypothesis`として登録し、別のサブエージェントによる補足・反証・確認レビューを
追記する。独立した確認レビューがあり、未解決点が空の場合だけ`confirmed`にする。
詳細な分析順序と確定条件は
[`ROOT_CAUSE_ANALYSIS.md`](ROOT_CAUSE_ANALYSIS.md)を正本とする。

確定した真因だけを`evaluation/remediations/ROOT-xxxxx.md`の対処タスクへ変換する。
対処タスクは評価候補と異なり、起点評価、変更対象・対象外、受入条件、回帰テスト、
実装コミットを保持する。`remediation-next`は次の`queued`タスクを返し、
`remediation-context`は実装担当へ渡す最小コンテキストを返す。

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
  root-causes/RC-xxxxx.md     # 真因仮説、独立レビュー、確定判断
  remediations/ROOT-xxxxx.md  # 確定した真因への設計対処タスク
  campaigns/*.yml            # 探索候補と優先度・coverage対象
  states/*.claim.yml         # 実行中claim（TTL付き）
  states/*.result.yml        # 完了した候補と評価ID・状態
  states/*.events.yml        # 状態遷移の追記型履歴（任意）
  coverage-index.yml         # 既存coverage JSONへの参照
  schema.yml                 # frontmatterと状態の検証規則
  legacy/EVAL_LOG.md         # 移行前ログの保全コピー
```

`EVAL_LOG.md` は手編集しない。評価ログのMarkdown生成コマンドはこれだけであり、
既定出力先はリポジトリルートの `EVAL_LOG.md` である。
`evaluation/EVAL_LOG.generated.md` は使用しない。

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
- 実Excel照合を使う評価では、必要な出力IDを列挙した`excelProbeIds`
- バグを確認した新規記録では `directCauseKey`, `causeKey`,
  `directFixStatus`, `rootFixStatus`。新方式で`rootFixStatus: fixed` の場合は
  `rootCauseId`、`rootFixTaskId`、`rootFixCommit`も必須。
- `horizontalAudit.confirmed`, `horizontalAudit.ruledOut`,
  `horizontalAudit.unresolved`
- EV-00385以降は`rootCauseAnalysis.status`, `directCause`, `designCause`,
  `confirmed`, `ruledOut`, `unresolved`, `decision`

### 期待値の根拠と検証ゲート

評価で「期待値と実際の動作が違う」と判断する前に、frontmatter の
`expectation` に根拠を記録する。`kind` は `spec`（仕様・公式文書）、
`excel`（Excel実機）、`hypothesis`（未確定の仮説）のいずれかである。
`statement` には具体的な期待動作を、`references` には仕様URL・文書または
実機ログとXL番号を記録し、`verification` は `pending`、`completed`、
`not-required` のいずれかにする。`spec` と `excel` は参照と検証完了が必須である。

`hypothesis` は最小再現、仕様確認、または実Excel照合を終えるまで
`bug-found`、`fixed`、`verified-no-bug` などの確定状態にしてはならない。
検証できない場合は `in-progress`、`needs-excel`、`blocked` のまま残し、
未確定の境界を `horizontalAudit.unresolved` に記録する。検証結果が想定を
否定した場合は期待値を修正し、推測に基づく実装修正を行わない。

過去の移行記録では `expectation` がない場合があるが、新規評価では必須の運用項目とする。

実Excel照合中は、状態を `needs-excel-probe` と `needs-excel` に分ける。
前者は `ExcelQueueVerification.bas` へのテスト作成待ち、後者はプローブ作成済みで
実機結果の反映待ちである。必要なXL番号は`excelProbeIds`を正本とする。評価ループは
未解決境界をすべて覆うプローブを追加し、`excel-sync`が`needs-excel`を返した同じ変更で
状態を更新する。`audit` は状態を推測して変更しない。

実機結果は、全プローブ終了時の`QUEUE_COMPLETE=True`、現在の
キューからインポートする`.bas`、`.cls`、`.frm`一式をLF改行へ正規化したSHA-256、
`excelProbeIds`の全出力が
そろった場合だけ同期済みと判定する。この三条件により、途中で停止した実行、古いVBAで
作った結果、一部のプローブだけを実行した結果を状態更新の根拠から除外する。
実機結果を反映した後は、
`verified-no-bug`、`known-limit`、`bug-found`、または修正後の `fixed` へ遷移して
完了を表す。実Excel確認の完了を別のサブステータスでは管理しない。

バグを修正した評価は、原因キー、回帰テスト、修正コミットを必ず記録する。
バグがない場合も `verified-no-bug`、実機照合待ちは `needs-excel`、恒久的制限は
`known-limit`、中断は `abandoned` または `blocked` として記録する。

使用可能な状態は次のとおりである。

`queued`、`claimed`、`in-progress`、`verified-no-bug`、`bug-found`、`fixed`、
`blocked`、`abandoned`、`known-limit`、`needs-excel-probe`、`needs-excel`、`retired`

### 完了日時の意味

状態が完了かどうかはスナップショットの`status`で判断できるため、新しい状態更新では
`completedAt`を生成しない。状態になった時刻は`events[].occurredAt`に記録し、これを
時系列集計の正本とする。`completedAt`は過去のスナップショットとの互換用に読み取る
だけであり、過去記録の評価日しかない場合も旧記録由来の日時として表示する。
CLIが作成する新しい結果スナップショットには`stateVersion: 1`を付け、対応する
`events`がない状態を`validate`で拒否する。旧形式のスナップショットはこの検査の対象外
であり、移行時にイベントを追加する。

## 現行CLIと役割

```text
npm run eval -- audit
npm run eval -- next --limit 1
npm run eval -- context <candidate-id>
npm run eval -- claim <candidate-id>
npm run eval -- release <candidate-id> <claim-token>
npm run eval -- record <evaluation-file>
npm run eval -- complete <candidate-id> <evaluation-id> <status> <claim-token>
npm run eval -- transition <candidate-id> <evaluation-id> <status> <claim-token>
npm run eval -- excel-sync <evaluation-id>
npm run eval -- remediation-next
npm run eval -- remediation-context <root-task-id>
npm run eval -- validate
npm run eval -- render EVAL_LOG.md
```

- `audit`: 記録、campaign、result、coverage参照を検証し、期限切れclaimを回収する。
- `next`: 実施可能な候補から、優先度とcoverage一致を考慮して選ぶ。
- `context`: 候補、関連評価、原因キー、横展開結果、coverageだけを出力する。
- `claim`: 候補を予約し、tokenを発行する。
- `release`: tokenを検証してclaimを解放する。
- `record`: 評価記録を冪等に保存する。同じIDの異なる内容は拒否する。
- `complete`: 候補と評価の対応、状態、修正コミットとテストを検証して結果を保存する。
  状態スナップショットには完了日時を重複保存せず、遷移日時はeventsへ記録する。
- `transition`: 待ち状態、`blocked`、`in-progress`、`bug-found`の再判定をclaim付きで行い、最新スナップショットを更新する。遷移前後はeventsに追記する。
- `excel-sync`: `excelProbeIds`、キューソース、実機結果を照合し、必要な待ち状態または`result-ready`をJSONで返す。状態は変更しない。
- `remediation-next`: 確定済み真因から作成した次の`queued`対処タスクを返す。
- `remediation-context`: 対処タスクと起点評価・FindingだけをJSONで返す。

実Excel照合結果の反映は、ログファイルを置くだけでは完了しない。先に`excel-sync`で
`result-ready`を確認し、runnerとの差異と未確定境界を本文の`unresolved`へ反映した後、
claim付きの`transition`で状態を更新する。`validate`は、キューから導出した待ち段階と
記録状態が違う場合、および`result-ready`なのに待ち状態のままの場合に失敗する。
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
`next` はその候補を自動的に除外する。次の評価では、直近の `directCauseKey` と
`causeKey`、横展開結果を再利用し、同じテストを理由なく繰り返さない。

### 原因分析と二段階の対処

直接修正だけで症状が消えても `rootFixStatus` は `unassessed` または `planned` のままに
する。真因への対処が必要なら、現在の発見評価のコミットを完了し、真因仮説の独立レビューを
終えてから`ROOT-xxxxx`を作成する。`remediation-next`で選んだタスクを`in-progress`へ
変更して実装し、受入条件・回帰テスト・横展開確認が完了した時点で`completed`とする。
その後、元の評価または関連Findingの`rootFixStatus`を`fixed`に更新し、`rootFixTaskId`と
コミットを紐付ける。真因が共通化不足でない場合も、変換、状態寿命、境界、評価順序、
エラー伝播、名前解決など同じ二層分析を行う。

### 状態履歴

`*.result.yml` は候補の現在状態を示すスナップショットであり、過去の状態を上書きする。
そのため、実機照合待ちなどを後から解消した場合にも時系列を失わないよう、完了または
再判定のたびに同じ候補の `*.events.yml` へイベントを追記する。イベントには
`candidateId`、`evaluationId`、`status`、`fromStatus`（2回目以降）、`occurredAt`
を保存し、最後のイベントはresultの状態・評価IDと一致しなければならない。

既存候補に履歴がないことは許容する。移行時に推測で過去イベントを作らず、次回の
状態更新から記録を開始する。`scripts/eval-report.mjs` はイベントファイルを読み取り、
レポートの状態遷移履歴に表示する。履歴がない過去状態の増減を推測してグラフへ追加する
ことはしない。

## 変更範囲に応じた検証

評価状態CLI、レポート、評価記録、ドキュメントだけを変更した場合は、次の短い検証で
十分とする。

```bash
npx tsx tests/tooling/evaluation-state.test.ts
npm run eval -- validate
node scripts/eval-report.mjs --html evaluation/EVAL_REPORT.html
npm run check:docs
```

`scripts/bg-test.sh` はエンジン、組み込み関数、Parser、拡張機能、共通テストハーネス
などランナーの実行経路を変更した場合に限って実行する。評価基盤だけの変更で全VBA
テストを繰り返さず、変更面に対応した検証結果を評価記録へ残す。短い検証で別領域の
失敗が判明した場合、または変更が複数領域に及ぶ場合は、全体テストへ昇格する。

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
