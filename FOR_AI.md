# FOR_AI.md — VBA Runner リファクタリング支援 AI 向けガイド

**VBA Runner** は **VBA実行環境 + リファクタリング支援ツール** です。
Excel 不要で VBA コードを実行・テスト・静的解析できます。

**このドキュメントはAI（LLM）を読者として想定しています。** 人間ユーザー向けの説明は含みません。

**このドキュメントの使い方:**
ユーザーから「このファイルの内容をもとにリファクタリングしてください」と指示されたら、
このドキュメントだけを読み、以下の**リファクタリングサイクル**を実行してください。
他のファイルを先読みする必要はありません。

---

## ⚠️ 基本ルール（違反禁止）

1. **VBAファイルの全文読み禁止** — `cat` / `Read` でファイル全体を読むことは禁止。
   常に `vba-analyzer` の行番号を使って必要な範囲だけ読む。
2. **計画単位で進める** — アナライザーの出力は候補一覧。全部やる指示ではない。Phase 2 で計画を立ててユーザーに承認を得てから実行する。計画内は自律実行し、計画完了後に次の計画をユーザーと合意する。
3. **テストなしで変更しない** — リファクタリング前に必ずテストを GREEN にしてから変更する。

---

## リファクタリングサイクル

```
                ┌─────────────────────────────────────────────┐
                │  計画が完了したら Phase 2（次の計画）へ戻る  │
                └──────────────────────┬──────────────────────┘
                                       ↑
Phase 1: 解析 → Phase 2: 計画 → Phase 3: 提案 → [ユーザー承認]
                                                        ↓
                              ┌─── Phase 4: リファクタリング ←─┐
                              │    Phase 5: テスト              │
                              │    Phase 6: 効果測定            │ 計画内の
                              │    Phase 7: レポート記録        │ 次のアイテムへ
                              │    コミット                     │
                              └────────────────────────────────┘
```

---

### Phase 1: 解析（Analysis）

**コマンド:**
```bash
# Step 1-a: 全体把握（全モジュール・全プロシージャを数十行に圧縮）
npx tsx test-libs/vba-analyzer.ts <対象ディレクトリ> --outline

# Step 1-b: 詳細解析（問題箇所・行番号付き）
npx tsx test-libs/vba-analyzer.ts <対象ディレクトリ>

# Step 1-c: コメントアウトコード確認
# コメントアウトコードが多いソースの場合無駄なトークン削減のためこの出力を使って
# 効果的に削除候補を判断、削除する
npx tsx test-libs/vba-analyzer.ts <対象ディレクトリ> --commented-code
```

**解析結果から読み取るもの:**

| 優先度 | 判断材料 | 意味 |
|---|---|---|
| 高 | `duplicateBlocks[].stmtCount × occurrences` が大きい | 重複解消のインパクトが最大 |
| 高 | `loopAnalyses[].extractionFeasibility` が低い（0〜0.5） | ループが抽出しやすい状態 |
| 中 | `maxNestDepth >= 5`（フラグ `nest=N`） | 深いネストの関数 |
| 中 | `lineCount >= 100`（フラグ `100L+`） | 巨大関数 |
| 低 | `excelObjectsUsed` がある | Excel依存でモック準備が必要 |

コードを読む必要がある場合は、アナライザーが示した**行番号の範囲だけ**読む:
```bash
# ✅ 正しい: 行番号指定
sed -n '79,115p' target.bas

# ❌ 禁止: 全文読み
cat target.bas
```

---

### Phase 2: 計画（Planning）

解析結果をもとに**複数件の優先順位付きリスト**を作る。
このリストが1回の「計画」単位。計画内の全件が完了したら次の計画へ移る。

**優先ルール（ユーザーが指定しない場合）:**
1. `duplicateBlocks` の `stmtCount × occurrences` が最大のもの
2. `excelObjectsUsed` が空かつ `ioSideEffectCount === 0`（純粋関数 → テストが書きやすい）
3. `loopAnalyses[].extractionFeasibility` が低いもの
4. 上記が同等なら `lineCount` が最大のもの

**計画フォーマット（内部で保持する。Phase 3 で提示する）:**
```
PLAN #N（<対象ディレクトリ名>）

#1 `<関数名>` (L<開始>-L<終了>, <行数>行)
   問題: <フラグ>
   変更: <1〜2文>
   期待: <指標> → <予測値>

#2 `<関数名>` (L<開始>-L<終了>, <行数>行)
   問題: <フラグ>
   変更: <1〜2文>
   期待: <指標> → <予測値>

（以下同様）

対象外:
- `<関数名>`: <理由>
```

---

### Phase 3: 提案（Proposal）

Phase 2 で作った計画をユーザーに提示し、承認を得る。
**承認なしに Phase 4 へ進んではいけない。**

```
## PLAN #N の提案

<Phase 2 の計画フォーマットをそのまま出力>

この順番で進めます。変更・追加があればお知らせください。
承認しますか？
```

ユーザーが承認したら → **Phase 4** へ（計画の #1 から順に実行）。
ユーザーが修正を求めたら → 計画を修正して再提示する。

---

### Phase 4: リファクタリング（Refactoring）

計画の現在のアイテムを実行する。ユーザーへの確認は不要（承認済みの計画に従う）。

**手順:**
1. アナライザーが示した行番号で対象コードを読む（`sed -n '<start>,<end>p'`）
2. 変更を加える
3. アナライザーを再実行して構文エラーがないことを確認する

**典型パターン:**

| 原因フィールド | 対処 |
|---|---|
| `assignmentBlocks` | 連続代入ブロックをデータ初期化関数に切り出す |
| `duplicateBlocks` | 重複する N 行を1つの関数に切り出し、両箇所から呼び出しに変える |
| `loopAnalyses` | 内側ループを `crossIterationVars` から ByRef パラメーターを推定して抽出 |
| `maxNestDepth >= 5` | 内側ループ・深い条件分岐をサブルーチンに切り出してガード節を使う |

---

### Phase 5: テスト（Testing）

**5-a. 既存テストの確認（変更前後の両方で必須）**

```bash
npx tsx <テストファイル>.ts
```

既存テストが RED になったら **計画を中断してユーザーに報告する。**

**5-b. 既存テストがない場合: リファクタリング前に記録する**

```typescript
import { VBARunner, assert } from '../../../test-libs/test-runner';
const vbaRunner = new VBARunner('sample/src/vba/MyModule.bas');
assert.strictEqual(vbaRunner.run('TargetFunction', [input1, input2]), expected, '説明');
```

`excelObjectsUsed` に列挙されたオブジェクトがある場合はモックが必要。
→ `docs/MOCK_GUIDE.md` の冒頭「Step 1: 対応表」でオブジェクト名を引く。

**5-c. 抽出した関数の単体テストを追加する（抽出直後に必須）**

統合テストの GREEN だけでは不十分。抽出した関数のエッジケース（空値・境界値・非数値等）を直接テストする。

```typescript
assert.strictEqual(vbaRunner.run('NewFunction', [edgeInput]), expected, 'エッジケースの説明');
```

全テスト（既存 + 新規）が GREEN になることを確認してから次へ進む。

---

### Phase 6: 効果測定（Measurement）

`--diff` モードを使って変更前後の指標を定量比較する。

```bash
# Phase 2 の計画立案前（リファクタリング開始前）にベースラインを保存しておく
npx tsx test-libs/vba-analyzer.ts <対象パス> --json > .vba-baseline.json

# Phase 6 で差分を表示（改善 ✅ / 悪化 ⚠ を自動判定）
npx tsx test-libs/vba-analyzer.ts <対象パス> --diff .vba-baseline.json
```

差分出力例:
```
=== TaskScheduler_Core.bas ===
  [ScanLockedRows]
    lineCount:     37 → 22  (-15) ✅
    maxNestDepth:   5 →  3  (-2)  ✅

=== ワークスペース集計 ===
  totalLines:      394 → 371  (-23)  ✅
  duplicateBlocks:   2 →   0  (-2)   ✅
```

この出力を Phase 7 のレポートにそのまま貼り付ける。ベースラインが存在しない場合はアナライザーをそのまま実行し、Phase 2 の計画から変更前の値を取得して手動で比較する。

---

### Phase 7: レポート記録（Report）

以下のフォーマットで出力する（ユーザーへの表示 + コミットメッセージの素材）。

```
## REPORT #N-<番号> `<関数名>`

### 変更内容
<何を変更したかを1〜2文>

### 定量的改善
| 指標 | 変更前 | 変更後 | 変化 |
|---|---|---|---|
| `<関数名>` 行数 | N行 | M行 | -X ✅ |
| `<関数名>` ネスト深さ | N | M | -X ✅ |
| フラグ | <変更前> | <変更後> | ✅ |
| 重複ブロック | N件 | M件 | -X ✅ |

### テスト
- 既存: N件 GREEN ✅
- 新規追加: N件 GREEN ✅（<エッジケースの説明>）

### 残り計画
- [ ] #<次番号> `<関数名>`
- [ ] #<その次> `<関数名>`
```

レポート出力後、**コミットする**（コミットメッセージはレポートの「変更内容」を使う）。

コミット後:
- **計画に次のアイテムがある** → Phase 4 に戻り次のアイテムを実行する
- **計画の全アイテムが完了** → Phase 2 に戻り次の計画を立てる（またはユーザーが終了を選択）

**重要:** アナライザーのフラグをゼロにすることが目的ではない。
「全フラグを消す」を目指すと過剰なリファクタリングになる。止め時の判断は常にユーザーに委ねる。

---

## ツールリファレンス

### vba-analyzer（静的解析）

```bash
npx tsx test-libs/vba-analyzer.ts <パス>                          # テキスト出力
npx tsx test-libs/vba-analyzer.ts <パス> --json                   # JSON出力（プログラム連携）
npx tsx test-libs/vba-analyzer.ts <パス> --json > baseline.json   # ベースライン保存
npx tsx test-libs/vba-analyzer.ts <パス> --diff baseline.json     # 差分表示（効果測定）
npx tsx test-libs/vba-analyzer.ts <パス> --diff baseline.json --json  # 差分をJSON出力
npx tsx test-libs/vba-analyzer.ts <パス> --outline                # AI向けコンパクト要約
npx tsx test-libs/vba-analyzer.ts <パス> --summary-only           # ワークスペース集計のみ
npx tsx test-libs/vba-analyzer.ts <パス> --commented-code         # コメントアウトコード候補のみ
```

**`--outline` の出力形式:**
```
[TaskScheduler_Core]  (210L)
  Function CalcDeadline  [nest=3]
  Function IsHoliday
  Private Sub UpdateStatus  ← 0 refs
[TaskScheduler]  (394L)
  Sub Main  [394L, nest=7, Excel×42]
```

**主要フィールド（プロシージャ単位）:**

| フィールド | 内容 |
|---|---|
| `name`, `kind`, `scope` | 名前・種別（Sub/Function/Property）・スコープ（public/private/friend） |
| `startLine`, `endLine`, `lineCount` | 行範囲と行数 |
| `maxNestDepth` | 最大ネスト深さ（5以上で `nest=N` フラグ） |
| `localDeclCount` | ローカル `Dim`/`Const` 宣言数 |
| `referenceCount` | 他プロシージャからの参照数（ワークスペース横断） |
| `assignmentBlocks[].shape` | 連続代入ブロックと形状: `const-decl` / `dim-decl` / `range-read` / `range-write` / `var-init` / `assign` / `mixed` |
| `excelAccessCount`, `excelAccessSamples[]` | Excel アクセス件数と行番号付きサンプル |
| `excelObjectsUsed[]` | モック必要候補オブジェクト（`Sheets`, `Range`, `Application` 等） |
| `hardcodedSheetCount`, `hardcodedAddressCount` | 固定シート名・固定セルアドレスの種類数 |
| `repeatedNumericLiterals[]` | 同じ数値リテラルが複数回登場（マジックナンバー候補） |
| `magicLiteralsInCalls[]` | 関数呼び出し引数のリテラル値（呼び出し先・引数インデックス付き） |
| `byRefAssignments[]` | ByRef パラメーターへの代入箇所 |
| `parameters[]`, `returnType` | パラメーター一覧（型・ByVal/ByRef）と戻り値型 |
| `ioSideEffectCount` | `MsgBox` / `InputBox` / `Debug.Print` の呼び出し回数 |

**主要フィールド（ファイル単位）:**

| フィールド | 内容 |
|---|---|
| `prefixClusters[]` | 接頭辞クラスター（例: `COL_×6` → Enum 化提案） |
| `gotoGraphs[]` | GoTo 解析（種別: `error_handler` / `cleanup` / `exit_loop` / `loop_skip` / `retry` / `other`、スコアと改善ヒント付き） |
| `loopAnalyses[].loops[].crossIterationVars[]` | ループ内変数の役割分析。詳細は下記 |
| `commentedCodeBlocks[]` | コメントアウトコード候補（`confidence` high/medium/low・`score`・`strippedContent` 付き） |
| `warnings[]` | パースエラー等の警告 |

**`crossIterationVars` の `role` と `extractionScore`:**

| `role` | 意味 | `extractionScore` |
|---|---|---|
| `loop_local` | 毎回書いてから読む。完全なローカル変数 | 0（そのまま抽出可能） |
| `accumulator` | `x = x + n` など前回値を累積 | 1（ByRef スカラーで渡す） |
| `state_flag` | Boolean 値または比較式のみ代入 | 1（ByRef で渡す） |
| `cross_iter` | 前の繰り返しで書いた値を次の繰り返しで読む | 1（ByRef で渡す） |
| `out_param` | ループ後に読まれる出力値 | 1 または 2（配列なら 2） |

`extractionFeasibility`（ループ単位の平均スコア）が低いほどループを関数に抽出しやすい。

**主要フィールド（ワークスペース横断）:**

| フィールド | 内容 |
|---|---|
| `entryPointCandidates[]` | Public・参照0（ヒューリスティック分類・理由付き） |
| `deadCodeCandidates[]` | Private・参照0（削除候補） |
| `callGraph[]` | `{from, fromFile, to, toFile}` プロシージャ間呼び出し |
| `duplicateBlocks[]` | 複数箇所に同一コードブロックが存在する重複 |

> **注意**: Excel ボタン・イベントハンドラー・`Application.Run` による呼び出しは静的解析の範囲外。
> 参照0のPublicプロシージャを「dead code」として削除しないこと。

---

### vba-formatter（フォーマッター）

**フォーマッターはAIのワークフローには含めない。リファクタリング後に人間が任意で実行するツール。**

```bash
npx tsx test-libs/vba-formatter.ts <ファイルまたはディレクトリ>            # stdout に出力
npx tsx test-libs/vba-formatter.ts <ファイルまたはディレクトリ> --check    # 差分チェック（exit 1）
npx tsx test-libs/vba-formatter.ts <ファイルまたはディレクトリ> --write    # 上書き
npx tsx test-libs/vba-formatter.ts <ファイルまたはディレクトリ> --indent-size=2
```

---

### VBA実行エンジン（テスト実行）

```typescript
// sample/tests/ts/MyFeature.test.ts
import { VBARunner, assert } from '../../../test-libs/test-runner';
const vbaRunner = new VBARunner('sample/src/vba/MyModule.bas');

// 引数ありの呼び出し
const result = vbaRunner.run('FunctionName', [arg1, arg2]);

// 引数なし / 式の評価
const val = vbaRunner.eval('SomeSub()');
```

```bash
npx tsx sample/tests/ts/MyFeature.test.ts
```

Excelオブジェクト（`ActiveSheet`, `Range`, `Cells` 等）はモック注入が必要。
→ `docs/MOCK_GUIDE.md` の冒頭「Step 1: 対応表」でオブジェクト名を引く。

**既知の制約:**
- Excel オブジェクトは自動でモック化されない
- `CreateObject("ADODB.Connection")` 等の COM オブジェクトはスタブ扱い
- ファイルI/O は `sandbox/` ディレクトリ以下に制限される

---

## 読む必要のないディレクトリ・ファイル

| パス | 理由 |
|---|---|
| `spec/` | MS-VBAL仕様書（4MB）。実行エンジン開発者向け |
| `src/lsp/` | VS Code LSP拡張機能の実装。リファクタリング作業と無関係 |
| `src/App.tsx` 他 React ファイル | Web UI |
| `dist/`, `node_modules/` | ビルド成果物・パッケージ |
| `TODO_SPEC.md` | VBA実行エンジンの仕様準拠TODO |
| `docs/internals/TYPE_SYSTEM_SPEC.md` | VBA実行エンジン内部型システムの仕様書 |
| `tests/lsp/`, `tests/engine/`, `tests/spec/` | VBA実行エンジン自体のテスト |

---

## サンプルコード（参照用）

| ファイル | 用途 |
|---|---|
| `sample/src/refactoring/TaskScheduler_v1.bas` | リファクタリング**前**の原典（394行の巨大関数） |
| `sample/src/refactoring/TaskScheduler_Core.bas` | リファクタリング**後**の純粋ロジック。テスト対象 |
| `sample/src/refactoring/TaskScheduler.bas` | リファクタリング後のメインルーチン（Excel依存部分） |
| `docs/REFACTORING_EXAMPLE.md` | TaskSchedulerのリファクタリング手順の記録 |
