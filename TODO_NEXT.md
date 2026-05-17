# 次期開発ロードマップ (Phase 2: VSCode IDE統合)

VBA Runner の付加価値は **VBAを実際に実行できること** と **レガシーVBAのリファクタリング支援** にあります。
この2点を軸に、VSCode拡張機能としての完成度を高めていきます。

> Phase 1（MS-VBAL 仕様書に列挙された構文要素・標準ライブラリ関数の実装）は完了。
> 仕様書本文に記載されたランタイム挙動の細部は [`TODO.md` の「VBA ランタイム挙動」](TODO.md#vba-ランタイム挙動) でトラッキング中。

---

## VBA Runner の付加価値

### VBA Runner なしでAI支援リファクタリングをする場合の限界

AIはVBAコードを読んでリファクタリング案を提案できる。しかし以下の問題がある：

- **提案が「仮説」のまま終わる** — AIの提案が正しいかどうかを確認するには Excel を開いてコードを貼り直す必要がある。イテレーションが重く、試行錯誤しにくい
- **大規模コードベースをAIに渡せない** — 数万行のVBAをそのまま貼り付けるとトークン上限に達する。どのファイルを渡すか人間が判断しなければならない
- **影響範囲がわからない** — 「この Sub を変えたらどこが壊れるか」をAIに問いかけても、コード全体を読んでいなければ答えられない
- **Dead codeがわからない** — 呼ばれていない関数を機械的に判定できない（この点は VBA Runner ありでも、Excel ボタン・イベントハンドラ・`Application.Run` による呼び出しは静的解析の範囲外）
- **テストがない状態でリファクタリングする** — 変更前の挙動を保存する手段がなく、リファクタリング後に壊れても気づけない

### VBA Runner がある場合

| 課題 | VBA Runner なし | VBA Runner あり |
|---|---|---|
| **提案の検証** | Excel で手動確認（重い） | その場で実行して即検証 |
| **大規模コード把握** | 人手でファイルを選んで貼り付け | AST から自動生成したアウトラインをAIに渡す |
| **影響範囲の把握** | AIが読んだ範囲での推測 | 参照解析による確定的な呼び出し元一覧 |
| **Dead code 検出** | 目視または grep | VBAコード内の参照数0プロシージャを列挙（Excel ボタン・イベントは検出外のため削除は要確認） |
| **リファクタリングの安全網** | なし（勘と経験） | 変更前に実行→スナップショットテストを自動生成 |
| **反復速度** | 変更→Excel貼り付け→手動実行 | 変更→テスト実行（数秒） |

### 本質的な差異：「仮説」から「検証済み」へ

AIによるリファクタリング提案は、VBA Runner がなければ **「おそらくこれで動くはず」** という仮説にとどまる。

VBA Runner があることで、AIが提案した変更を **その場で実行して検証** できる。
結果として、AIと人間のやり取りが「提案を読んで手動確認する」ループから
「提案→実行→結果確認→次の提案」という **高速な反復サイクル** に変わる。

さらに、実行できることで **スナップショットテストの自動生成** が可能になる。
これは他のVBA解析ツールにはない機能で、「リファクタリング前の挙動を保存する」
安全網を低コストで敷けるため、大胆なリファクタリングへの心理的障壁を下げる。

---

## Step 1 — 繋ぐだけで使えるもの（実装済み・未接続）

実装とテストは存在するが、extension.ts に登録されていないだけの機能。

- [ ] **Diagnostics（構文エラーの破線表示）**
  - [x] Parser が `ast.diagnostics[]` にエラーを収集（エラー耐性パース）。テスト: `lsp-diagnostics.test.ts`
  - [ ] DiagnosticsProvider クラスの作成（`ast.diagnostics` → `vscode.Diagnostic[]` 変換）
  - [ ] extension.ts への登録（`createDiagnosticCollection`、ドキュメント変更時に更新）
- [ ] **DocumentSymbol（アウトライン表示）**
  - [x] SymbolProvider 実装済み。テスト: `lsp-symbol-provider.test.ts`
  - [ ] extension.ts への登録（`registerDocumentSymbolProvider`）
- [ ] **DAP（デバッガ）の VSCode 接続**
  - [x] ステップ実行・ブレークポイント・変数表示 実装済み。テスト: `lsp-debugger.test.ts`, `lsp-debug-adapter.test.ts`
  - [ ] extension.ts への登録（`registerDebugAdapterDescriptorFactory`）

---

## Step 2 — リファクタリングに必須の LSP 機能

これがないと「安全に変更できる」とは言えない。シンボルテーブルが既にあるので実装コストは比較的低い。

- [ ] **Find All References（Shift+F12）**
  - 「この Sub はどこから呼ばれているか」を一覧表示。削除・改名の影響範囲を把握するために必須
  - `textDocument/references` の実装と extension.ts への登録
- [ ] **Rename Symbol（F2）**
  - Sub / 変数の名前を全ファイル横断で一括変更
  - `textDocument/rename` の実装と extension.ts への登録

---

## Step 3 — 「実行できる」固有の価値を活かす機能

他のVBAエディタとの差別化。VBAを実行できるVBA Runner ならではの機能。

- [ ] **Code Lens（各 Sub/Function 上のインライン情報）**
  ```
  ▶ Run  |  3 references  |  未テスト
  Function CalcTotal(a As Long, b As Long) As Long
  ```
  - `▶ Run` ボタン：クリックで即実行し結果を Output Panel に表示
  - `N references`：参照元数を表示（`0 references` は削除候補の可能性。ただし Excel ボタン・イベントハンドラ・`Application.Run` からの呼び出しは静的解析では検出できないため、削除前に要確認）
  - `未テスト / テスト済み`：`Test_` プロシージャの有無で表示
- [ ] **Dead code 検出 と エントリーポイント候補の分離**（analyzer 側は実装済み、LSP 化が残）
  - VBAコード内からの参照が 0 件のプロシージャを2つに分類する：
    - **Private** + 参照0 → **Dead code 候補**（削除検討対象）
    - **Public** + 参照0 → **エントリーポイント候補**（Excel ボタン・イベント・`Application.Run` から呼ばれている可能性が高く、削除すべきではない）
  - 命名規則ヒューリスティック：`Workbook_*` / `Worksheet_*` / `Auto_*` はイベントハンドラ、`Test_*` はテストプロシージャ、それ以外の長めの Public Sub は「マクロボタンの入口」と推定
  - LSP 接続として、エントリーポイント候補は「実行ボタン」を Code Lens で表示する起点に活用
  - **注意**: 静的解析では検出できない呼び出し元（`Application.Run` の文字列呼び出し等）が依然として残るため、最終判断は人間が行う
- [ ] **安全な部分抽出支援（Extract Sub/Function）**
  レガシーVBAで頻出する「単調だが長大なブロック」を安全に切り出すための支援。
  抽出・分割は壊れやすく、実行して検証できる VBA Runner の価値が直接発揮される。
  - **抽出候補の検出**: ASTで連続する代入文のブロック（例: セルから変数への一括読み込み、フィールドの初期化など）を検出してハイライト。複雑度は低いが行数が多く、抽出によりトークン削減効果が高い
  - **抽出前スナップショット**: 対象プロシージャを実行して入出力を記録
  - **抽出後の自動検証**: 抽出後に同じ入力で再実行し、挙動が変わっていないことを確認
  - **巨大ファイルの早期分割支援**: モジュールの責務境界（処理の塊の変わり目）をASTから推定し、分割位置を提案。大きなファイルを早めに分割することで、それ以降のAI支援作業のコンテキストサイズを削減できる
- [ ] **VBA to TypeScript トランスパイラ**
  - 構築済みの AST から TypeScript コードを自動生成する Generator の実装
  - リファクタリングの最終ゴール：VBA からのモダン言語移行支援

---

## Step 4 — AI支援によるリファクタリング加速

レガシーVBAの現実：数万行が数十ファイルに分散し、どこから手をつければいいかわからない。
VBA Runner は AST・実行・参照解析をすべて持っているため、AI（Claude等）との連携で
**「何をどう直すべきか」を効率よく把握する** ことができる。

- [ ] **Workspace Outline 生成（AI向けコンテキスト圧縮）**
  - ASTから全モジュール・全プロシージャのシグネチャ一覧を生成し、コンパクトなテキストとして出力
  - 例: `VBA: Copy AI Context` コマンド → クリップボードに構造化サマリーをコピー
  - 効果: 50,000行のVBAをAIに丸ごと貼る代わりに、2,000トークン程度のサマリーで全体像を把握させられる
  - 出力形式イメージ:
    ```
    [Module: TaskScheduler_Core]
      Function CalcDeadline(startDate, days) → Date
      Function IsHoliday(d) → Boolean
      Sub UpdateStatus(taskId, status)   ← 0 refs from other modules
    [Module: TaskScheduler]
      Sub Main()  ← calls: CalcDeadline, IsHoliday, Sheets("Tasks").Range(...)
    ```

- [ ] **Call Graph 解析 / Excel依存マップ**
  - プロシージャ間の呼び出し関係をツリー構造で可視化
  - `Sheets(...)` / `Range(...)` / `Application` への依存箇所を自動特定し、ロジックとI/Oの境界を明確化
  - **テスト可能な純粋関数の自動識別**：Excel依存なし・グローバル変数非参照のプロシージャを候補として列挙
  - AIへの活用: 「このモジュールのどの関数から先にテストを書くべきか」の判断材料を構造化データで提供

- [ ] **Code Smell 定量化レポート**
  静的解析で問題箇所を数値化してリスト化。AIにそのまま渡してリファクタリング提案を求めるワークフロー。
  - Dead code（VBAコード内からの参照数0のプロシージャ。Excel ボタン・イベント・`Application.Run` からの呼び出しは検出不可のため「削除候補」として扱う）
  - 長大関数（行数ランキング上位N件）
  - 深いネスト（インデントレベル閾値超え）
  - `On Error Resume Next` の多用（バグを隠蔽しているリスク箇所）
  - マジックナンバー（説明なしのリテラル数値）
  - グローバル変数の多用（スコープ汚染）

- [ ] **スナップショットテスト自動生成**
  - 既存プロシージャを実際に実行して入力→出力のペアを記録し、回帰テストコードとして自動生成
  - 「リファクタリング前の挙動を保存する」安全網を低コストで敷ける
  - 他のVBAツールにはできない、**実行できるからこそ可能な機能**

- [ ] **ピンポイントコンテキスト生成（カーソル位置の関数）**
  - カーソル位置のプロシージャについて「呼び出し元一覧」「呼び出し先一覧」「参照しているグローバル変数」を構造化してクリップボードにコピーするコマンド
  - → そのままAIチャットに貼り付けて「この関数をどう分割すべきか」「テストを書いて」と相談できる
- [ ] **トークン削減を目的とした早期分割の提案**
  - 巨大ファイル（例: 2000行超）を検出し、「ここで分割すると後続のAI作業で X% コンテキストが削減される」という見積もりを提示
  - 代入文の連続など「大きいが単純なブロック」を自動検出してリストアップ。`セルから変数への一括読み込み` などは抽出コストが低く、抽出後のファイルサイズ削減効果が高い典型例
  - 分割・抽出は壊れやすいため、実行検証（スナップショット）とセットで提案する

---

---

## リファクタリング支援ツール `test-libs/vba-analyzer.ts`

### 目的

Step 4（AI支援リファクタリング）の機能群を **段階的に試作・改善するための支援基盤** として整備。
レガシーVBAを解析してリファクタリング候補を機械的に列挙する CLI ツール。
LSP 機能として VSCode に組み込む前の **プロトタイプの場** であり、ここで使い物になった機能を
順次 LSP 化していく。

### 使い方

```bash
./node_modules/.bin/esbuild test-libs/vba-analyzer.ts --bundle --outfile=test-libs/vba-analyzer.cjs --platform=node
node test-libs/vba-analyzer.cjs <file-or-dir>           # テキスト形式（人間が読む用）
node test-libs/vba-analyzer.cjs <file-or-dir> --json    # JSON 形式（プログラム/AI 連携用）
```

### 現在できること

`sample/src/vba_legacy/TaskScheduler_v1.vba` および `sample/src/vba/` での実証結果を例に。

| 機能 | 検出結果の例 |
|---|---|
| プロシージャごとの行数・ネスト深さ・ローカル変数数 | v1: 394行・最大ネスト8段・Dim 66件 |
| フラグ表示（LARGE / DEEP_NEST / MANY_LOCALS / EXCEL_HEAVY） | v1: 4つすべて点灯 |
| 連続代入ブロックの検出（形状付き） | v1: 6箇所、最大25件（L96-L129, shape:mostly-dim-decl） |
| Excel I/O アクセス箇所のカウントと **行番号付き** サンプル | v1: L77/L78/L79... を含む25件 |
| 繰り返し現れる数値リテラル（マジックナンバー候補） | v1: 0.25(×4), 0.5(×3) |
| **エントリーポイント候補の一覧化**（Public・参照0）| v1: `AutoScheduleTasks` を "likely button/macro entry point" として明示 |
| **Excel モック必要箇所の一覧化**（リファクタリング前準備）| v1: `AutoScheduleTasks` には `ActiveSheet, Application, Cells, Columns, Range, Rows` のモックが必要と提示 |
| **Dead code 候補**（Private・参照0）| v1 では該当なし。Private プロシージャのみ対象 |
| **接頭辞クラスタ検出**（UDT/Enum 抽出候補）| v1: `COL_`×6, `CONFIG_`×4, `ROW_`×3 を検出 |
| **呼び出しグラフ（Call Graph）**| `AutoScheduleTasks → 21関数` のクロスファイル依存関係を可視化 |
| **Workspace Outline**（`--outline`）| 全モジュール・プロシージャをコンパクトに要約。AI向けコンテキスト圧縮に活用 |
| ワークスペース横断のクロスファイル参照解析 | プロシージャの呼び出し関係を計数 |
| ファイル/ディレクトリ両対応、テキスト/JSON/outline 出力 | `.bas/.cls/.frm` を再帰的に処理 |
| `--summary-only` / `--outline` / `--json` モード | 用途に応じた出力形式を選択可能 |

### 不足機能（このツールで埋めていく TODO）

- [x] **Expression レベルの位置情報（loc）の付与**
  - `src/engine/parser.ts` の `parsePrimary()` / binary / unary メソッドに `loc` を追加
  - MemberExpression, CallExpression, BinaryExpression, UnaryExpression すべてに `loc` が付くようになった
  - Excel I/O サンプルが `L?` → `L77` などの実際の行番号に変わった

- [x] **代入ブロックの形状クラスタリング**
  - `assignmentBlocks` の各エントリに `shape` プロパティを追加
  - 分類: `const-decl`, `dim-decl`, `var-init`, `range-read`, `range-write`, `set-obj`, `assign`, `mixed`
  - 多数派が 50% 超なら `mostly-X`、それ以外は `mixed` と表示

- [ ] **データフロー解析（Def-Use チェーン）**
  - 「この行範囲を抽出する場合、引数として渡すべき変数」を機械的に算出
  - 「この行範囲で書き換えている変数」（ByRef 引数候補）を機械的に算出
  - 抽出リファクタリング（Extract Function）の安全性を担保する基礎
  - 出力: 行範囲を指定すると `inputs[]` / `outputs[]` / `local[]` を返す API を追加

- [x] **識別子の接頭辞クラスタ検出**
  - `COL_*` × 6件 / `CONFIG_*` × 4件 / `ROW_*` × 3件 を検出（TaskScheduler_v1.bas での実証）
  - `prefixClusters: [{ prefix, members[], suggestion }]` として出力

- [ ] **コメント領域からの責務境界推定**
  - `' Phase 1: Scan Locked Rows` `' =====` のような領域コメントを「人間が書いた責務境界の宣言」と見なし、**分割位置の第一候補** として提案
  - AST だけでは見えない設計意図がコメントに現れている
  - Lexer はコメントを捨てているため、コメントトークンの保持が前段に必要
  - 出力: `regionMarkers: [{ line, label, depth }]`

- [ ] **リファクタリング前後の対応追跡（Refactoring Drift Tracking）**
  - 2ファイルを与えると「v1 のどの行範囲が現状のどの抽出関数になったか」を機械的に追跡
  - 副次効果：抽出漏れ・対応外行が検出できる
  - 出力: `--diff <file_before> <file_after>` モードを追加

- [ ] **Excel API スタブモード（実行系の拡張）**
  - 解析ではなく Evaluator 側の話だが、解析→実行検証の流れで詰まる箇所
  - Analyzer が「このプロシージャには `ActiveSheet, Application, Cells, ...` のモックが必要」と一覧化できるようになったので、その一覧をもとに `ActiveSheet` / `Range` / `Cells` を自動でモック化するモードを設ける
  - 「とにかく実行してみたい」段階のためのデフォルト挙動。値を返すスタブで十分
  - これがないと「リファクタリング前の挙動を保存する」スナップショット生成の入り口で詰まる
  - 連携イメージ: `vba-analyzer --json` の `excelMockTargets` を読み込んで、必要なモックを自動構築してから実行

- [x] **呼び出しグラフ（Call Graph）の生成**
  - `callGraph: [{ from, fromFile, to, toFile }]` として出力（JSON / テキスト両対応）
  - `AutoScheduleTasks → 21関数` のような依存関係が可視化された

- [x] **Workspace Outline 形式での出力**
  - `--outline` モードで全モジュール・全プロシージャをコンパクトに出力
  - `[Excel×N]` `[nest=N]` `← 0 refs` フラグ付きで問題箇所が一目でわかる

- [x] **重複ブロック検出（AST N-gram マッチング）**
  - プロシージャ内・プロシージャ間で同一パターンが繰り返すステートメント列を検出
  - 変数名は `$ID` に正規化済みのため、名前が違っても構造が同じなら一致
  - スライディングウィンドウのノイズ除去（同一 proc 内のオーバーラップを先着優先で排除）
  - より長いパターンに包含される短いパターンを除去（maximal match フィルタ）
  - 出力: `duplicateBlocks: [{ stmtCount, length, shape, occurrences[] }]`
  - TaskScheduler_v1.bas での実証: Range読み込み4連パターン・Application設定リストア・グリッド読み込みループなど

---

## リファクタリングスキルの設計（完成後の利用イメージ）

Step 2〜4 の機能が揃った後、以下のスキルとプロンプトで使えるようになることを目指す。

### スキル定義（`.claude/commands/refactor-vba.md` の概要）

```
/refactor-vba <対象ファイルまたはディレクトリ>
```

**手順:**
1. **全体把握** — Workspace Outline を生成してモジュール構成・問題箇所を列挙
2. **対象選定** — 優先順位を提案し、ユーザーの承認を得る
3. **安全網** — スナップショットテストを作成してコミット（ベースライン確保）
4. **リファクタリング** — Excel I/O とロジックの分離、長大関数の分割など
5. **検証** — テスト全通過を確認して段階的にコミット

---

### プロンプト例

**全体把握:**
```
このVBAコードのアウトラインを出してください。
モジュールごとに「何をしているか」と主なプロシージャ一覧を
300行以内でまとめてください。
```

**依存関係の調査:**
```
TaskScheduler.bas の中で Sheets / Range / Application に
依存しているプロシージャを列挙してください。
依存していない（= そのままテスト可能な）プロシージャも教えてください。
```

**問題箇所の洗い出し:**
```
以下の観点でコードを分析してリストアップしてください：
- 50行を超えるプロシージャ
- On Error Resume Next を使っているプロシージャ
- どこからも呼ばれていないプロシージャ（Dead code の候補。Excel ボタン・イベントハンドラ・`Application.Run` からの呼び出しは検出できないため削除前に要確認）
```

**安全網の作成（VBA Runner 固有：実行できるから可能）:**
```
CalcDeadline 関数を代表的な入力パターンで実際に実行して、
結果をスナップショットテストとして書いてください。
リファクタリング後の回帰テストとして使います。
```

**単調な代入ブロックの抽出（トークン削減と可読性向上）:**
```
ProcessMonthlyReport の冒頭にあるセルから変数への代入が50行続いています。
この部分を LoadReportData() という Sub に抽出してください。
抽出前に実行してスナップショットを取り、抽出後も同じ結果になることを確認してください。
```

**巨大ファイルの早期分割:**
```
TaskScheduler.bas は3000行あります。
後続のAI作業のコンテキストを減らすために、まず責務ごとにファイルを分割したいです。
分割の境界になりそな箇所をASTから分析して提案してください。
分割後も動作が変わらないことをテストで確認してからコミットしてください。
```

**Excel I/O とロジックの分離:**
```
CalculateSalesTotal を以下のように分離してください：
1. Excel I/O：セルの読み書き（そのまま残す）
2. 純粋ロジック：計算だけを行う Function として抽出
抽出した Function にユニットテストも作成してください。
```

**ワンショット（全工程を一括依頼する場合）:**
```
legacy/TaskScheduler.bas をリファクタリングしてください。
VBA Runner を使って以下の順で進めてください：

1. 問題箇所（長大関数・Excel依存混在・Dead code）を列挙
2. テストしやすい関数を3件選んで優先順位を提案
3. 承認後、スナップショットテストを作成
4. ロジックと Excel I/O を分離するリファクタリングを実施
5. テストで検証してコミット

各ステップで確認してから次に進んでください。
```

---

## 完了済みの機能

### VSCode 拡張機能（extension.ts に接続済み）
- [x] Tolerant Parsing — Lexer 列番号、Parser `loc` 位置情報、エラー耐性パース
- [x] Hover — シンボルのシグネチャ表示。テスト: `lsp-hover.test.ts`
- [x] Go to Definition（F12）— シンボル宣言位置へジャンプ。テスト: `lsp-definition.test.ts`
- [x] Completion — 標準関数・変数補完。テスト: `lsp-completion.test.ts`
- [x] Test Explorer — `Test_` プロシージャの自動検出と実行。テスト: `lsp-test-discovery.test.ts`, `lsp-test-runner.test.ts`

### テスト基盤（TypeScript テストランナー）
- [x] Spy / Mock API（`MsgBox`, `Shell` 等の副作用検証）。テスト: `spy-mock-api.test.ts`
- [x] Time Mocking（`Now` / `Date` の固定）。テスト: `time-mocking.test.ts`
- [x] In-Memory FS（`MemoryFileSystem`）の基礎実装

### Extension インフラ
- [x] `src/extension.ts`（Desktop 用 Node.js 版）
- [x] `language-configuration.json`, `syntaxes/vba.tmLanguage.json`
- [x] `package.json`（VSCode extension metadata）、`.vscodeignore`

---

## 保留中・将来課題

### パッケージング
- [ ] VSCode Marketplace への公開（アイコン整備、`.vsix` ビルド検証）
- [ ] Web Extension 化（evaluator が Node.js `path` に依存しており、先にブラウザ対応が必要）

### Web UI（デモサイト）の改善
- [ ] `Dir` 関数の完全実装（ディレクトリ列挙）
- [ ] `Kill` ワイルドカード対応（例: `Kill "*.txt"`）
- [ ] コールスタックの出力・`Erl` 関数サポート

---

**最終ビジョン**:
VBA開発者が VBE を開かずに VSCode だけで、LSP 補完を受けながらコードを書き、
参照・リネームで安全にリファクタリングし、テストを即座に実行し、
ステップデバッグで挙動を確認できる「モダンなVBA開発環境」を実現する。

さらに、数万行のレガシーVBAに対して、コードベース全体をAIに効率よく把握させ、
「どこから手をつけるか」「どう分割するか」をAIと対話しながら進められる
**AI支援リファクタリングの起点** となることを目指す。
