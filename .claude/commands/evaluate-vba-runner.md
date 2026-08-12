# vba-runner 使い勝手評価スキル

評価状態の正本は `docs/internals/EVALUATION_STORAGE.md` に定義する構造化記録へ
段階移行する。移行前は `EVAL_LOG.md` を読み取り専用の旧記録として扱い、移行後は
`evaluation/evaluations/*.md`、キャンペーン状態、カバレッジ参照を正本とする。
`EVAL_LOG.md` は正本から生成するビューであり、手編集しない。

真因分析の方法・証拠基準・ステータス判定は
`docs/internals/ROOT_CAUSE_ANALYSIS.md`を正本とする。真因分析サブエージェントには
この文書を先に読ませ、内容をコピーせず、その形式で報告させる。

毎回の評価で全履歴をサブエージェントへ渡してはならない。`eval audit` と
`eval next` が返す候補、直近の関連評価、同じ原因キーの横展開結果、対象カバレッジ
だけをプロンプトへ含める。全履歴の読み込みは移行・監査・収束判定に限定する。

サブエージェントに vba-runner を実際に使わせてサンプル VBA プログラムを作らせ、
使い勝手・改善点・バグをレポートさせる。新規ユーザー視点での「生の使用感」を
得るための手順。

## 手順

### 0. 評価状態を監査して候補をclaimする

`evaluation` の監査を先に実行し、staleなclaimや参照切れを解消する。
その後、カバレッジ未通過箇所、既知の原因キー、実機未照合境界を考慮して
`eval next --limit 1` で候補を1件選び、`eval claim` で予約する。claimが返す秘密tokenを
保持し、サブエージェントには
`eval context <candidate-id>` の出力だけを渡す。

30分待機後は、候補選択の前に`eval audit`で全体状態を確認し、残るRCA・横展開・
実機待ち・通常候補の優先順位から次の作業を決める。この判断はキャンペーンに依存
しない`evaluation/checkpoints/EVALUATION_LOOP.yml`へ、監査結果・理由・判断結果
（既存候補の`claim`または新規候補の`create-candidate`）・次候補とともに
追記する。EV本文とキャンペーン定義には運用チェックポイントを記録しない。

移行前の一時運用では、従来どおり `EVAL_LOG.md` を読み取り専用で参照する。
ただし、候補選定に不要な全履歴をサブエージェントへ渡さない。

#### 継続評価手法の選定

未実施領域が空でも評価を終了しない。構造化キャンペーンから、前回と異なる
ファザーの生成シード・構文組み合わせ・入力境界、またはミューテーションの
対象ファイル・変異演算子・テスト単位を選ぶ。同じ対象を再実行する場合は、
回帰確認である理由を記録する。

実施後は評価記録とキャンペーン状態を更新する。確認済み、除外、未解決を分けて
記録し、1回のファザー生成やミューテーション実行を手法全体の完了とは扱わない。
中断時は状態を消さず、`abandoned` または `blocked` として次回に回収できるようにする。

対象の優先順位は、利用可能な最新カバレッジの未通過分岐・行を第一の手がかりにする。
ファザーは未通過のパーサー／評価器経路へ到達する入力を、ミューテーションは未通過
経路を含むファイルの境界変異を優先する。カバレッジが古ければ
`./scripts/coverage.sh --report` で既存データだけを再集計し、データがない場合や大きな変更後に限り全計測を
行う。選定時には、使用したカバレッジJSONを `coverage-index.yml` のスナップショットとして
参照し、評価記録からリンクする。既存coverage JSONの形式や生成物は変更しない。
カバレッジは仕様適合やバグの不存在を証明しないため、実機照合結果と既知の制限も
併せて確認する。

### 1. サブエージェントを起動する

Agent ツール（`subagent_type: general-purpose`）を1つ起動する。サブエージェントは
この会話の文脈を持たないため、プロンプトは自己完結させること。プロンプトには
以下を含める。

- **vba-runner の説明**: `test-libs/test-runner.ts` がライブラリ本体（`VBARunner` /
  `assert` を export）であり、`tsx` でビルド不要に直接 import できること
  （`import { VBARunner, assert } from './test-libs/test-runner';`）
- 公式の利用者向けドキュメント `build/runner/README.md` を先に読ませる
  （これは実際に npm 公開されているパッケージの README そのもの＝「ユーザーが見る
  ドキュメント」）
- **候補の評価コンテキストを読ませる**。`eval next` が出力した対象、関連する評価記録、
  同じ原因キーの横展開結果、既知の制限だけを渡す。移行前は必要な範囲に限って
  `EVAL_LOG.md`（`/Users/koji/src/github.com/jca02266/vba-runner/EVAL_LOG.md`）を参照させ、
  毎回全文をプロンプトへ複製しない
- **オリジナルの、ある程度複雑な VBA サンプルプログラムを1つ設計・実装させる**
  （ドメインは自由だが、評価済みドメインとは被らないようにすること）。
  最低限以下を含めること（ただし当該評価回の「注力テーマ」で追加条件があればそちら優先）:
  - Class モジュール1つ以上
  - Collection または配列
  - 文字列処理
  - `On Error` によるエラーハンドリング
  - ループ・条件分岐などの制御構文
- サンプルの `.bas`/`.cls` ファイルと TypeScript ドライバーコードは、**このリポジトリの
  追跡対象外**（例: `/tmp/vba-runner-eval/`）に置かせる。リポジトリ内のファイルは
  変更させない
- `npx tsx` でリポジトリルートから実行させ、動くまで（または明確にバグ・制限に
  当たるまで）反復させる。回避せず、当たった問題は最小再現コードまで絞り込ませる
- **`TODO_SPEC.md` / `TODO.md` / git log を事前に読ませない**こと（既知の課題リストに
  引っ張られない、フレッシュな視点での評価を得るため）
- 最後に、最終メッセージとして以下を日本語または英語で 800 語程度以内に
  まとめて返すよう指示する:
  - 作ったサンプルの概要
  - 使い勝手評価（表形式や箇条書きで良い。何が滑らかで、何が引っかかったか）
  - 実際に当たった具体的な改善提案（一般論的なwishlistではなく、体験に紐づくもの）
  - 発見したバグ（最小再現コード・期待動作と実際の動作・エラーメッセージを添えて）
  - スクラッチファイルを削除したかどうか

より広い網羅性が欲しい場合は、観点を分けて（例:「エラー処理」「ファイル入出力」
「Excel スタブ」等）複数の Agent を**1メッセージ内で並列起動**してもよい。

### 2. 報告内容を自分で検証する

サブエージェントの報告は**そのまま信用しない**。特にバグ報告は、報告された
最小再現コードを自分で `Bash`（`npx tsx -e "..."` など）で実際に実行し、同じ
症状が再現するか確認する。

再現できたバグは、`src/engine/evaluator.ts` 等の関連ソースを読んで根本原因を
簡潔に特定する（可能な範囲で）。再現できなかった報告は「未確認」として報告から
外すか、その旨を明記する。

### 2.3 真因分析と対処タスクのゲート

#### 期待値の根拠を先に確定する

期待値だけを想定してバグ判定してはならない。評価記録のfrontmatterに
`expectation`を追加し、`kind`（`spec` / `excel` / `hypothesis`）、具体的な
`statement`、`references`（仕様のURL・文書、または実機ログとXL番号）、
`verification`（`pending` / `completed` / `not-required`）を記録する。
`spec` と `excel` は根拠の参照と検証完了が必須である。`hypothesis` は
最小再現、仕様確認、または実Excel照合を終えるまで確定状態にしてはならず、
その間は `in-progress`、`needs-excel`、または `blocked` として扱う。
検証結果が想定を否定した場合は期待値を修正し、推測に基づく実装修正を行わない。
`npm run eval -- validate` はこのゲートを検証する。

実Excel照合は `needs-excel-probe` と `needs-excel` の2状態で管理する。
前者は `ExcelQueueVerification.bas` へのプローブ作成待ち、後者はプローブ作成済みで
実機結果の反映待ちである。必要なXL番号はfrontmatterの`excelProbeIds`へ列挙する。
評価ループは未解決境界をすべて覆うプローブを追加した後に`eval excel-sync`を実行し、
その判定が`needs-excel`になった同じ変更で状態を更新する。`eval audit` は状態を推測して
変更しない。`eval excel-sync`は実機結果のハッシュに加えて、準備スタンプと現在ソースの
一致も報告する。スタンプが不一致の場合は`prepare-excel-vba.sh`を再実行するまで
`needs-excel`への遷移やWindows依頼を行わない。
実機結果を反映したら、
`verified-no-bug`、`known-limit`、`bug-found`、または修正後の `fixed`へ遷移する。
同じ領域の実機候補は最大5件までまとめて登録・準備し、対象に応じたWindowsコマンドを
1回だけ実行する。通常キューは`eval-excel.cmd`、FormatとRadixのマトリックスは
両方を連続実行する`eval-matrix.cmd`を使う。結果ファイルから各EVを個別に照合・遷移する。

バグを再現し、横展開調査を終えたら、修正を始める前に真因分析サブエージェントを
起動する。症状の最小再現、関連ソース、横展開結果だけを渡し、次を独立に報告させる。

分析サブエージェントは、まず `docs/internals/ROOT_CAUSE_ANALYSIS.md` を読み、そこに
定義された順序で期待値、最初の誤状態、直接原因、設計上の真因、確認済み・除外・
未解決経路、対処案を報告する。最初の報告は必ず真因仮説として
`evaluation/root-causes/RC-xxxxx.md`へ登録し、分析担当自身に確定させない。

- 直接原因（`directCauseKey`）と、そこへ至る設計・状態・変換・ディスパッチ上の
  真因（`causeKey`）を分ける
- 同じ真因を持つ確認済み経路、除外経路、未解決経路
- 共通化、局所修正、仕様上の制限、TODO化のどれが妥当かという根拠
- 修正する場合の影響範囲と必要な回帰テスト

分析サブエージェントには追跡ファイルを変更させない。仮説登録後、別のサブエージェントに
同じ入力と仮説を渡し、補足、反証、代替仮説、説明できない経路をレビューさせる。
レビューで未解決点が増えた場合は、別担当による最小調査を続ける。独立した確認レビューが
あり、未解決点が空になるまで`confirmed`へ遷移させず、実装タスクも作らない。

報告を受けたら、まず評価記録と
Findingに2つの原因キー、`directFixStatus`、`rootFixStatus`を反映する。さらに評価記録の
`rootCauseAnalysis`へ、`status`、`directCause`、`designCause`、
`confirmed`、`ruledOut`、`unresolved`、`decision`を個別に保存する。
`horizontalAudit`は挙動の横展開、`rootCauseAnalysis`は原因の分析と対処判断なので、
片方で代用してはならない。真因が未確定ならバグを確定せず、
追加の最小検証または`unresolved`として残す。

過去の評価を参照するときは、`rootCauseProcedureVersion`を必ず確認する。
`0`の`rootCauseAnalysis.status: confirmed`は旧方式の記録としてのみ扱い、独立レビュー済み
の真因根拠へ昇格させない。`1`は`rootCauseId`のRC記録、同じ手順バージョン、独立
`confirm`レビュー、空の`unresolved`まで確認する。方式不明なら、新しい仮説として登録する。

真因が確定した場合は、`evaluation/remediations/ROOT-xxxxx.md`に、起点評価、確定した
`rootCauseId`、対象・対象外、受入条件、必要なテストを持つ真因対処タスクを登録する。
評価キャンペーン候補を真因対処タスクの代用にしてはならない。直接修正のコミット完了後、
このタスクを`in-progress`へ遷移して別の実装タスクを起動する。実装タスクは分析で定めた
範囲だけを変更し、元の回帰テストと横展開で確認対象になった経路のテストを追加・更新する。
共通化不足は真因の一例にすぎないため、変換の不一致、状態寿命、境界条件、評価順序、
エラー伝播、名前解決なども同じ手順で検討する。実装タスク完了後はrootが差分、回帰テスト、
真因への対処が一致していることを再検証する。直接修正と真因修正が別コミット・別候補なら、
真因対処タスクIDとコミットを記録してから`rootFixStatus: fixed`へ遷移する。直接修正だけで
評価を`fixed`にする場合も、`rootFixStatus`を別途明示して未対処を隠さない。

### 2.5. バグ修正後にレグレッションテストを必ず追加する

**評価で発見したバグを修正したら、必ずレグレッションテストを追加すること。**
これを省くと同じバグが再発したときに気付けない。

- LSP 機能のバグ（フォーマッター、コードレンズ、補完など）→ `tests/lsp/` の対応するテストファイルに追加
- エンジン（evaluator/parser/lexer）のバグ → `tests/spec/` の対応するテストファイルに追加
- 新規テストファイルが必要な場合は既存テストのインポート・スタイルに従って作成する
- テストは修正が正しく機能することを確認し、修正前は失敗・修正後は通過することを `npx tsx` で確認してからコミットする

### 3. 構造化評価記録を保存する

今回の評価結果を `evaluation/evaluations/EV-xxxxx.md` に保存する。

- frontmatterの状態、優先度、キャンペーン、coverage参照を更新
- 本文に再現手順、結果、原因、横展開のconfirmed/ruled-out/unresolvedを記録
- バグがあれば原因キー、回帰テスト、修正コミットを関連付ける
- 実Excel待ち、恒久制限、未実施を別状態として保存する
- `eval complete <candidate-id> <evaluation-id> <status> <claim-token>` で候補と評価結果を関連付ける
- `eval validate` と `eval render` を実行し、生成された `EVAL_LOG.md` を更新する
- 新規EVは`evaluationRecordVersion: 2`を必須とする。`eval record`はこの項目がない
  新規記録を拒否する。終端状態への`complete`もv2本文の5見出し、期待値・結果、判定を
  検証してから状態を保存する。終端の`## 判定`には`判定状態: <status>`を置き、YAMLの
  `status`と一致させる。

### 3.5 実Excel照合結果を状態へ反映する

`tests/excel/queue/*.result` などに実機ログを追加しただけでは、評価状態やレポートは変わらない。
必要なテストIDは評価記録の`excelProbeIds`を正本とし、次の順序で必ず状態へ反映する。

1. `npm run eval -- excel-sync <evaluation-id>`を実行する。`result-ready`でなければ比較を
   始めない。欠落ID、未完了、ソースハッシュ不一致はコマンド出力に従って解消する。
2. 実機出力とrunnerの同一入力を比較し、期待値・エラー番号・未確定の仕様差を評価本文の
   `unresolved` から削除または更新する。ログが候補の一部しか覆わない場合は、未確認の
   境界を `unresolved` に残す。
3. 差異が確定した場合は通常のバグ修正手順へ戻り、実機一致で差異がない場合は
   `verified-no-bug`、実装しない仕様差は `known-limit` として記録する。
4. claimを取得し、`npm run eval -- transition <candidate-id> <evaluation-id> <status> <token>`
   を実行してresultとeventsを更新する。評価記録のfrontmatter、本文、実機ログの対応を
   同じ変更で保存する。
5. `npm run eval -- validate` と `npm run eval -- render EVAL_LOG.md` を実行し、レポートの
   `needs-excel` 件数が実際の未確定境界と一致することを確認する。

確定後に期待値の誤りや重複が判明した場合は、BUGを`retired`へ変更して
`retiredReason`を記録し、EV本文を訂正する。claim取得後、次でイベントを残したまま
EVを巻き戻す。

```bash
npm run eval -- claim <candidate-id> --rollback
npm run eval -- rollback <candidate-id> <evaluation-id> <status> <claim-token> <reason>
```

`status`は`in-progress`、`verified-no-bug`、`known-limit`、`blocked`のいずれかとする。
CLIは以前の終端状態、v2本文、claim、理由を検証し、`rollbackFrom`と`rollbackReason`を
イベントへ追記する。挙動を改めて評価する場合は巻き戻しではなく、新しい候補とEVを作成する。

通常キューの`ExcelQueueVerification.result`は、全プローブ終了を示す
`QUEUE_COMPLETE=True`と、現在のキューからインポートする`.bas`、`.cls`、`.frm`一式を
LFへ正規化したSHA-256が一致して初めて同期済みになる。マトリックス実行では
`FormatMatrix.result`の`FORMAT_MATRIX_COMPLETE=True`と`RadixMatrix.result`の
`RADIX_MATRIX_COMPLETE=True`をそれぞれ同じソースハッシュ検証付きで受理する。
実機ログの存在だけを根拠に`needs-excel`件数を減らしてはならない。また、評価記録を
更新せずにログだけをコミットしてはならない。`validate`は導出された待ち段階と記録状態の
不一致、および同期済み結果を待ち状態のまま残すことを拒否する。

キューの`.bas`、`.cls`、`.frm`が更新された場合、開発側で
`tests/excel/queue/prepare-excel-vba.sh`を実行する。このコマンドが
`empty_with_macro.xlsm`を`t.xlsm`へコピーし、`vba-extractor import`を実行する。
同時に`t.xlsm.source.sha256`を作成するので、`t.xlsm`と同じ場所へコピーする。
Windows側の`eval-excel.cmd`または`eval-matrix.cmd`はExcel起動前にこのスタンプと現在のVBAソースを照合し、
準備漏れがあれば停止する。
Windowsへ依頼する前に開発側でも同じスタンプ照合を実行する。照合が成功した場合だけ、
プローブ未作成の`needs-excel-probe`を実機結果待ちの`needs-excel`へ遷移させ、評価記録と
準備済みブックを同じ変更として扱う。
生成された`t.xlsm`をWindowsへ渡し、Windows側の`eval-excel.cmd`は通常キューを、
`eval-matrix.cmd`はFormatとRadixの2つのマトリックスを準備済みブックで実行し、各結果を
UTF-8へ変換するだけとする。Windows側にNode環境を要求せず、ソース更新ごとに準備コマンドを
再実行してから実機へ渡す。

新しい実Excel照合候補を登録するときは、未使用の`XL-xxx`を予約して評価記録の
`excelProbeIds`へ列挙する。まだプローブがなければ`needs-excel-probe`とする。その後、通常
プローブは`tests/excel/queue/ExcelQueueVerification.bas`へ、Format/Radixマトリックスは
各専用モジュールへ追加し、対応するPublic runnerから呼び出す。`excel-sync`が`needs-excel`を返したら、評価記録と
resultをclaim付きで同じ状態へ遷移する。クラスや
フォームが必要な場合は同じディレクトリに`.cls`/`.frm`を追加する。評価記録の`tests`には
このVBAソース、追加ソース、XL番号をすべて列挙する。スクラッチで動作確認しただけの
候補を実機照合キューとして登録してはならない。

### 4. ユーザーに報告する

検証済みの内容（サンプル概要・使い勝手評価・改善提案・検証済みバグと根本原因）を
日本語で簡潔にユーザーへ報告する。バグ修正が必要そうな場合でも、無断で実装には
進まず、修正してよいかユーザーに確認する。

## コミット後の待機

評価コミット後の30分待機は、会話を占有するsleepや1分ごとの状態報告で実施しない。
`scripts/eval-wait.sh`で待機プロセスをバックグラウンド起動し、IDだけを控える。
待機中は他の会話や作業を継続できる。完了後に一度だけ`status`を確認し、次の評価へ進む。
開始時に絶対的な`resume_at`を保存するため、待機プロセスが途中で終了しても、再開時刻を
過ぎていれば再待機せず完了として扱う。

```bash
wait_id="eval-$(date +%Y%m%d-%H%M%S)"
./scripts/eval-wait.sh start 1800 "$wait_id"
./scripts/eval-wait.sh status "$wait_id"
```

`status`が`state=done`になるまで次の評価を開始しない。待機を取り消す必要がある場合は、
IDを指定して停止する。後で同じ待機を継続する場合は`resume`を使い、別の待機を始める場合だけ
新しいIDを発行する。

```bash
./scripts/eval-wait.sh stop "$wait_id"
./scripts/eval-wait.sh status "$wait_id"
```

`state=interrupted`は待機未完了を意味するため、コミット後30分の条件を満たしたとは
扱わない。ただし、再開時に`resume_at`を過ぎていれば、次のコマンドで完了へ更新できる。

```bash
./scripts/eval-wait.sh resume "$wait_id"
```

`resume`は残り時間だけを待ち、期限を過ぎていれば即時に`state=done`へ遷移する。
待機の状態ファイルはOSの一時ディレクトリに保存し、リポジトリへ追加しない。
