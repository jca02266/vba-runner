# 次期開発ロードマップ (Phase 2: VS Code IDE統合)

VBA Runner の付加価値は **VBAを実際に実行できること** と **レガシーVBAのリファクタリング支援** にあります。
この2点を軸に、VS Code拡張機能としての完成度を高めていきます。

> Phase 1（MS-VBAL 仕様書に列挙された構文要素・標準ライブラリ関数の実装）は完了。
> 仕様書本文に記載されたランタイム挙動の細部は [`TODO.md` の「VBA ランタイム挙動」](TODO.md#vba-ランタイム挙動) でトラッキング中。

---

## VBA Runner の付加価値

### VBA Runner なしでAI支援リファクタリングをする場合の限界

AIはVBAコードを読んでリファクタリング案を提案できる。しかし以下の問題がある：

- **提案が「仮説」のまま終わる** — AIの提案が正しいかどうかを確認するには Excel を開いてコードを貼り直す必要がある。イテレーションが重く、試行錯誤しにくい
- **大規模コードベースをAIに渡せない** — 数万行のVBAをそのまま貼り付けるとトークン上限に達する。どのファイルを渡すか人間が判断しなければならない
- **影響範囲がわからない** —「この Sub を変えたらどこが壊れるか」をAIに問いかけても、コード全体を読んでいなければ答えられない
- **Dead codeがわからない** — 呼ばれていない関数を機械的に判定できない（この点は VBA Runner ありでも、Excel ボタン・イベントハンドラー・`Application.Run` による呼び出しは静的解析の範囲外）
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

- [x] **Diagnostics（構文エラーの破線表示）**
  - [x] Parser が `ast.diagnostics[]` にエラーを収集（エラー耐性パース）。テスト: `lsp-diagnostics.test.ts`
  - [x] `LSPServer.getDiagnostics()` 実装（`ast.diagnostics` → LSP Diagnostic 変換）。テスト: `server-diagnostics.test.ts`
  - [x] extension.ts への登録（`createDiagnosticCollection`、ドキュメント open/change/close 時に更新）
- [x] **DocumentSymbol（アウトライン表示）**
  - [x] SymbolProvider 実装済み。テスト: `lsp-symbol-provider.test.ts`
  - [x] extension.ts への登録（`registerDocumentSymbolProvider`）
- [x] **DAP（デバッガー）の VS Code 接続**
  - [x] ステップ実行・ブレークポイント・変数表示・実装済み。テスト: `lsp-debugger.test.ts`, `lsp-debug-adapter.test.ts`
  - [x] `src/lsp/vscode-debug-adapter.ts` 作成（`VBADebugAdapterFactory` + `VBAInlineDebugAdapter`）
  - [x] extension.ts への登録（`registerDebugAdapterDescriptorFactory`）
  - [x] package.json に `contributes.debuggers` エントリ追加（launch設定テンプレート含む）
- [ ] **DAP ステップ実行の本実装**
  現状の `debugger.ts` は DAP プロトコルの接続枠組みのみで、`stepOver` / `stepInto` / `stepOut` / `continue` はスタブ（state 変更のみ）。
  F10 を押してもカーソル位置が動かず、変数パネルも実値を返さない。

  **必要な変更:**
  - **Evaluator のジェネレーター化**: `function*` で各文の実行境界で `yield` し、外部から一文ずつ進められるようにする
  - **現在行トラッキング**: 実行中の Statement の `loc.start.line` を保持し、`stackTrace` レスポンスに反映する
  - **変数パネルの実装**: 実行中の `Environment` スコープから変数名・値・型を取得して返す
  - **ブレークポイント停止**: 各文実行前にブレークポイントリストと照合し、一致したら `stopped(breakpoint)` イベントを送出
  - **stepOver vs stepInto の分岐**: `stepOver` はサブルーチン呼び出しを単一ステップとして扱い、`stepInto` は呼び出し先の先頭で止まる

  **実装アプローチ（推奨）:**
  ```typescript
  // Evaluator をジェネレーターに変換
  function* executeStatements(stmts: Statement[], env: Environment) {
      for (const stmt of stmts) {
          yield stmt.loc?.start.line; // ← ここで止まれる
          yield* executeStatement(stmt, env);
      }
  }
  ```

  **難度:** ★★★（Evaluator 全体の再設計が必要）

---

## Step 2 — リファクタリングに必須の LSP 機能

これがないと「安全に変更できる」とは言えない。シンボルテーブルがすでにあるので実装コストは比較的低い。

- [x] **Find All References（Shift+F12）**
  - 「この Sub はどこから呼ばれているか」を一覧表示。削除・改名の影響範囲を把握するために必須
  - `src/lsp/references-provider.ts` 実装（テキストベース全単語検索、コメント・文字列リテラル除外）
  - extension.ts への登録（`registerReferenceProvider`）。テスト: `lsp-references.test.ts`
- [x] **Rename Symbol（F2）**
  - Sub / 変数の名前を全ファイル横断で一括変更
  - `src/lsp/rename-provider.ts` 実装（References を利用してテキスト編集リストを生成）
  - extension.ts への登録（`registerRenameProvider`）。テスト: `lsp-references.test.ts`

---

## Step 3 — 「実行できる」固有の価値を活かす機能

他のVBAエディターとの差別化。VBAを実行できるVBA Runner ならではの機能。

- [x] **Code Lens（各 Sub/Function 上のインライン情報）**
  ```
  ▶ Run  |  3 references  |  未テスト
  Function CalcTotal(a As Long, b As Long) As Long
  ```
  - `▶ Run`：引数なし（Optional のみ含む）のプロシージャに表示、クリックで即実行→Output Panel
  - `N references`：参照元数を表示（`⚠ 0 references` は Private Dead code 候補）
  - `未テスト / ✓ テスト済み`：`Test_*` プロシージャから参照されるかで判定
  - `src/lsp/code-lens-provider.ts`。テスト: `lsp-code-lens.test.ts`（9件）
  - `vba-runner.runProcedure` / `findReferences` / `generateTest` / `goToTest` コマンド登録
- [x] **Dead code 検出 と エントリーポイント候補の分離**
  - **Private** + 参照0 → Diagnostic Warning `'名前' は参照されていません (Private Dead code 候補)`
  - **Public** + 参照0 → Code Lens で `0 references` 表示のみ（エントリーポイント候補のため警告なし）
  - 戻り値代入（`FuncName = value`）や再帰呼び出しをプロシージャ自身の本体内と判定して除外
  - `CodeLensProvider.getDeadCodeWarnings()` → `getDiagnostics()` にマージ
  - **注意**: 静的解析では `Application.Run` 等の動的呼び出しは検出不可。最終判断は人間が行う
  - テスト: `lsp-dead-code.test.ts`（5件）
- [ ] **`__mocks__` ディレクトリによるモック注入**
  Excel 依存コードをモックに差し替えて VS Code の `▶ Run` から実行できるようにする仕組み。
  実行できるという固有の価値を Excel 依存コードにも広げる。

  **ファイル構成（2つの形式をサポート）:**

  *単一ファイル形式*（小規模・シンプル）:
  ```
  src/vba/
  ├── TaskScheduler.bas
  ├── TaskScheduler_runner.bas
  ├── __mocks__.bas             ← VBA モック 1 ファイルにまとめる
  └── __mocks__.ts              ← TypeScript モック 1 ファイルにまとめる
  ```

  *ディレクトリ形式*（大規模・ファイル分割）:
  ```
  src/vba/
  ├── TaskScheduler.bas
  ├── TaskScheduler_runner.bas
  └── __mocks__/
      ├── ExcelObjects.bas      ← VBA モック（Sheets / Range / Application のスタブ）
      ├── ExcelObjects.ts       ← TypeScript モック（同名も可）
      └── MsgBox.ts             ← MsgBox / InputBox を Debug.Print に差し替え
  ```

  両形式が同時に存在する場合はどちらも読み込む。

  **読み込み優先順序（`▶ Run` 実行時）:**
  ```
  ① AssertHelper（常に注入）
  ② __mocks__.ts / __mocks__/*.ts → esbuild でバンドル → Evaluator に built-in として登録
  ③ __mocks__.bas / __mocks__/*.bas/.cls → 結合コードの先頭に追加
  ④ 同ディレクトリの本番 .bas/.cls
  ```

  **VBA モック（`.bas`/`.cls`）の例:**
  ```vb
  ' __mocks__/ExcelObjects.bas
  Function Sheets(name)
      Set Sheets = New MockSheet
  End Function

  Function ActiveSheet()
      Set ActiveSheet = New MockSheet
  End Function
  ```

  **TypeScript モック（`.ts`）の例:**
  ```typescript
  // __mocks__/MsgBox.ts
  // default export で { 関数名: 実装 } を返す
  export default {
      MsgBox: (prompt: string) => {
          console.log(`[MsgBox] ${prompt}`);
          return 1; // vbOK
      },
      InputBox: (_prompt: string) => "mock-input",
  };
  ```

  TypeScript モックは esbuild でバンドルして `require()` し、`Evaluator` の組み込み関数テーブルに登録する。
  これにより「Node.js の fs を使った高度なスタブ」「外部 API を返すモック」なども書ける。

  **Evaluator 側の変更（必要な拡張）:**
  - `Evaluator` コンストラクターに `builtinOverrides?: Record<string, Function>` を追加
  - TypeScript モックの export をそこに渡す
  - 組み込み関数テーブルの検索で `builtinOverrides` を優先する

  **実装上の注意点:**
  - VBA モック内で同名 Sub/Function を先に定義しても、後から読み込む本番コードで上書きされる可能性がある。
    エバリュエーターの「後勝ちか先勝ちか」を確認してから実装順を決める
  - 組み込み `MsgBox` / `Debug.Print` は Evaluator 内部に実装があるため、TypeScript モック経由（`builtinOverrides`）でないと上書きできない
  - `__mocks__/` は `.vscodeignore` 対象外（本番ビルドには含まない想定だが、VS Code 拡張の配布には不要）

  **将来拡張（vba-analyzer との連携）:**
  - `vba-analyzer --json` の `excelMockTargets` を読んで `__mocks__/ExcelObjects.bas` ひな形を自動生成するコマンド
  - 例: `▶ Generate Mocks` Code Lens → `vba-analyzer` 出力をもとに最小限のスタブファイルを生成

  **作業順序（実施時）:**
  - [ ] Evaluator に `builtinOverrides` オプションを追加
  - [ ] extension.ts の `runProcedure` で `__mocks__/` を検出・ロード
  - [ ] TypeScript モック: esbuild バンドル → `require()` → `builtinOverrides` 登録
  - [ ] VBA モック: 結合コードの先頭に追加（本番コードが後）、定義衝突の動作を確認
  - [ ] LSP.md に `__mocks__` 仕様を追記

- [ ] **Signature Help（引数入力支援）**
  関数呼び出しの `(` 内を入力中に、現在カーソルがある引数の名前・型・説明をツールチップ表示する。
  VBA エディター（VBE）の「クイックヒント」相当の機能。キーワード引数（`name:=value` 形式）の入力補助にも使える。

  **実装難度は情報ソース別に3段階に分かれる:**

  | 対象 | 難度 | 取得手段 |
  |---|---|---|
  | ユーザー定義 Sub/Function | ★☆☆ 低 | AST の `ProcedureDeclaration.parameters` をそのまま利用 |
  | 組み込み VBA 関数（MsgBox, Left 等） | ★★☆ 中 | 200〜400 関数分の静的 JSON 定義ファイルを用意 |
  | Excel/COM 定数（xlUp, xlDown, VbMsgBoxStyle 等） | ★★★ 高 | 型ライブラリ（.tlb）または `@types/excel` などからインポート |

  **推奨実装順:**
  1. **ユーザー定義プロシージャのみで先行リリース**（高効果・低コスト）
     - カーソルが `(` 〜 `)` 内にあるかを検出
     - カンマを数えてアクティブな引数インデックスを特定
     - `signatureHelp` レスポンスでシグネチャと強調引数を返す
  2. 組み込み VBA 関数の静的 DB を追加（`src/lsp/builtin-signatures.json`）
  3. Excel enum 定数は外部定義 JSON をロードする仕組みを先に設計し、DB は段階的に充実させる

  **注意:** Excel 型ライブラリの定数は数千件あり、完全な型推論（引数型 → enum候補の絞り込み）には型推論機構が必要。
  まず「引数名と型名の表示」から始め、「型名に応じた候補一覧」は後回しにするのが現実的。

- [ ] **Code Actions（右クリック → Refactor メニュー）**
  `vscode.languages.registerCodeActionsProvider` で `CodeActionKind.Refactor` を登録し、
  右クリックメニューの「Refactor...」に以下の項目を追加する。

  | アクション | トリガー | 難度 |
  |---|---|---|
  | **✅ 変数の導入**（Introduce Variable） | ✅ 実装済み。複数置換時にユーザーに確認を求める機能付き（全て置換/一つずつ確認/置換しない） | ★☆☆ |
  | **定数の抽出**（Extract Constant） | リテラル値を選択 → プロシージャ先頭に `Const name = <値>` を挿入し全箇所を置換 | ★☆☆ |
  | **変数のインライン化**（Inline Variable） | 変数名にカーソル → `Dim name = <式>` を削除して全参照箇所に式を展開 | ★★☆ |
  | **関数の抽出**（Extract Sub/Function） | 複数行を選択 → 新 Sub/Function に切り出し、元の位置に呼び出しを挿入。ByRef パラメーターはクロスイテレーション変数解析から推定 | ★★★ |
  | **With ブロックの導入**（Introduce With） | 同一オブジェクトの連続アクセス（`obj.A = / obj.B =`）を選択 → `With obj ... End With` でラップ | ★★☆ |
  | **マジックナンバーの定数化** | リテラル数値を選択（Code Smell 検出と連携）→ Extract Constant と同じ動作 | ★☆☆ |

  **実装上のポイント:**
  - `Inline Variable` は参照数が 1 の場合のみ安全に適用可（複数参照は副作用の可能性あり）
  - `Extract Sub/Function` は vba-analyzer の `crossIterationVars` を流用してパラメーター候補を自動推定
  - 変数名は `showInputBox` でユーザーに入力させる（デフォルト候補を出す）

- [ ] **Source Actions（右クリック → ソースアクション メニュー）**
  `vscode.languages.registerCodeActionsProvider` で `CodeActionKind.Source` を登録し、
  右クリックメニューの「ソースアクション...」に以下の項目を追加する。
  カーソル位置に関係なくファイル全体を対象とする定型操作。

  | アクション | 説明 | 難度 |
  |---|---|---|
  | **Organize Declarations** | `Option Explicit` をファイル先頭に追加し、各プロシージャ内の未宣言変数を検出して `Dim` 文を自動挿入する | ★★★ |
  | **Remove Unused Variables** | 参照数 0 の `Dim` 宣言を削除する（References プロバイダーを利用） | ★★☆ |

  **Organize Declarations の想定ユースケース:**

  `Option Explicit` がないレガシーコードは変数を `Dim` なしで使っていることが多い。
  このアクションを適用すると：
  1. `Option Explicit` をファイル先頭に挿入
  2. 各プロシージャを AST で走査し、`Dim` 宣言のない識別子（代入先・使用箇所から推定）を収集
  3. 各プロシージャの先頭に `Dim <名前> As Variant`（型推定できる場合は推定型）を挿入

  **実装上のポイント:**
  - 未宣言変数の検出: AST の代入文左辺・`For` ループ変数など「値を書き込む箇所」を起点にする
  - 型推定: 右辺が数値リテラルなら `Long`、文字列なら `String`、推定不可なら `Variant`
  - モジュールレベル変数（`m` プレフィックス慣習など）は各プロシージャではなくファイル先頭付近に挿入
  - `Remove Unused Variables` は `ReferencesProvider` の参照数 0 判定を流用。ただし `Dim` 文が複数変数を宣言している場合（`Dim a, b As Integer`）は行単位ではなく変数単位で削除するため、トークン再組み立てが必要

- [ ] **VBA 固有の落とし穴に対する Diagnostics 警告**
  VBA 初心者・移行者が踏みやすい言語仕様の罠を静的解析で検出し、`Warning` または `Information` 診断として表示する。
  各ルールは設定でオン/オフできるようにする。

  **型・宣言系:**

  | 対象パターン | 説明 | 重大度 |
  |---|---|---|
  | `Dim a, b As Long` | `a` は `As Long` が適用されず `Variant` になる。各変数に型を明示すべき | Warning |
  | パラメーターに `ByVal`/`ByRef` なし | VBA のデフォルトは `ByRef`。意図しない呼び出し元への副作用が起きやすい | Warning |
  | `Integer` 型の使用 | VBA では `Long` の方が高速（32bit/64bit 共通）。`Integer` は後方互換のみ | Information |

  **エラーハンドリング系:**

  | 対象パターン | 説明 | 重大度 |
  |---|---|---|
  | `On Error Resume Next` 後に `Err.Number` チェックなし | エラーを握り潰している可能性 | Warning |
  | `On Error GoTo 0` の欠落 | `On Error Resume Next` を使ったあとエラー抑制が解除されずに続く | Information |
  | エラーハンドラーで `Resume` / `Resume Next` なし | `GoTo ErrHandler` で飛んでも `Resume` がないと制御が戻らない | Warning |

  **非推奨・旧構文系:**

  | 対象パターン | 説明 | 重大度 |
  |---|---|---|
  | `While...Wend` | 古い構文。`Do While...Loop` を推奨 | Information |
  | `Global` キーワード | 旧構文。`Public` を推奨 | Information |
  | `GoTo`（エラーハンドラー以外） | スパゲッティ化の原因 | Warning |

  **コードスメル系:**

  | 対象パターン | 説明 | 重大度 |
  |---|---|---|
  | 未使用パラメーター | 宣言されているが本体内で参照されていない引数 | Warning |
  | 変数シャドウイング | モジュールレベル変数と同名のローカル変数 | Warning |
  | `Select Case` に `Case Else` なし | 想定外の値を無言でスルー | Information |
  | 同一リテラルが 3 回以上 | マジックナンバー（定数化候補） | Information |

  **Excel 操作系（脆弱パターン）:**

  | 対象パターン | 説明 | 重大度 |
  |---|---|---|
  | `Sheets(1)` / `Sheets(2)` などの数値インデックス | シート順変更で壊れる。名前指定を推奨 | Warning |
  | `ActiveSheet` / `ActiveWorkbook` の使用 | 何が選択されているか依存。明示的な参照を推奨 | Information |
  | `ScreenUpdating = False` の戻し忘れ | エラー発生時に `True` に戻されない（フロー解析が必要） | Warning |

  **実装上のポイント:**
  - `Dim a, b As Long` の検出: `VariableDeclaration` ノードで宣言リストを走査し、末尾以外に型指定がない変数を警告。Quick Fix で `Dim a As Long, b As Long` に展開
  - `ByVal`/`ByRef` 未明示: `ProcedureDeclaration.parameters` を走査し、修飾子がない引数に警告
  - 未使用パラメーター: `ReferencesProvider` を流用してプロシージャ本体内の参照数 0 を検出
  - 変数シャドウイング: モジュールレベルのシンボルテーブルとローカルスコープの名前を照合
  - `ScreenUpdating = False` の戻し忘れ: プロシージャ内の代入を追跡し、`False` 設定後に `True` に戻す経路がない場合に警告（フロー解析）
  - Excel 操作系は Excel 依存コード以外では無意味なため、`.cls` ファイルまたはモジュール先頭の宣言からクラス種別を推定してから適用

- [ ] **Cyclomatic Complexity の可視化**
  - **Code Lens**: 各 Sub/Function の宣言行に CC スコアを表示（CC ≤ 10: 表示なし、11〜15: `CC: 12`、> 15: `CC: 18 ⚠️`）
  - **Diagnostics**: CC > 閾値（デフォルト 15）の関数に `Information` / `Warning` 診断を出力
  - **Decoration（任意）**: `If` / `For` 等の各分岐行末尾に `/* +1 */` のゴーストテキストで貢献度を表示

- [ ] **安全な部分抽出支援（Extract Sub/Function）**
  レガシーVBAで頻出する「単調だが長大なブロック」を安全に切り出すための支援。
  抽出・分割は壊れやすく、実行して検証できる VBA Runner の価値が直接発揮される。
  - **抽出候補の検出**: ASTで連続する代入文のブロック（例: セルから変数への一括読み込み、フィールドの初期化など）を検出してハイライト。複雑度は低いが行数が多く、抽出によりトークン削減効果が高い
  - **抽出前スナップショット**: 対象プロシージャを実行して入出力を記録
  - **抽出後の自動検証**: 抽出後に同じ入力で再実行し、挙動が変わっていないことを確認
  - **巨大ファイルの早期分割支援**: モジュールの責務境界（処理の塊の変わり目）をASTから推定し、分割位置を提案。大きなファイルを早めに分割することで、それ以降のAI支援作業のコンテキストサイズを削減できる
- [x] **✅ コールグラフの可視化**
  ✅ 実装済み。プロシージャ間の呼び出し関係を VS Code の Webview パネルに Mermaid.js で表示。
  レガシー VBA の「どの関数がどこから呼ばれているか」の把握を視覚的に支援。

  **表示内容:**

  | ノード種別 | 色例 | 説明 |
  |---|---|---|
  | エントリーポイント候補（Public・参照0） | 青 | ボタン・イベントから呼ばれている可能性 |
  | 通常の Sub/Function | 白 | 呼び出し元・呼び出し先あり |
  | Dead code（Private・参照0） | 赤 | 削除候補 |
  | Excel 依存あり | オレンジ枠 | Sheets/Range/Application を直接参照 |

  **操作:**
  - ノードクリック → 定義へジャンプ（`Go to Definition`）
  - ノードを右クリック →「ここからの呼び出し元のみ表示」「ここへの呼び出し先のみ表示」でフィルタリング
  - 「Mermaid としてコピー」ボタン → ドキュメント貼り付け用テキストをクリップボードへ

  **起動方法:**
  - コマンドパレット: `VBA: Show Call Graph`
  - コードレンズまたは右クリックメニュー: `Show in Call Graph`（カーソル位置の関数を中心に表示）

  **出力イメージ（Mermaid）:**
  ```
  graph TD
      Main --> CalcDeadline
      Main --> IsHoliday
      Main --> UpdateStatus
      CalcDeadline --> IsWeekend
      IsHoliday --> IsWeekend
      UpdateStatus:::dead
      classDef dead fill:#f88
  ```

  **実装方針:**

  | フェーズ | 内容 | 難度 |
  |---|---|---|
  | 1. グラフ構築 | AST を走査して `CallExpression` を収集し `caller → callee[]` の Map を生成 | ★☆☆ |
  | 2. Webview パネル | `vscode.window.createWebviewPanel` で HTML パネルを開き、Mermaid.js で描画 | ★★☆ |
  | 3. インタラクション | ノードクリック時に Webview → Extension へ `postMessage` し `vscode.commands.executeCommand('editor.action.goToDeclaration')` | ★★☆ |
  | 4. Excel 依存マーク | `Sheets` / `Range` / `Application` / `ActiveSheet` などのキーワードをノードに付加 | ★☆☆ |
  | 5. フィルタリング | 選択ノードの祖先・子孫のみ表示する部分グラフ生成 | ★★☆ |

  **実装上の注意点:**
  - `CallExpression` の呼び出し先が文字列（`Application.Run "ProcName"`）の場合は静的解析不可。警告表示のみ
  - 再帰呼び出しはループとして表示（Mermaid の `graph TD` は循環を許容）
  - ファイルをまたぐ呼び出し（同ディレクトリの複数 `.bas`）はフェーズ 2 以降で対応

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
  - AIへの活用:「このモジュールのどの関数から先にテストを書くべきか」の判断材料を構造化データで提供

- [ ] **vba-analyzer 差分モード（リファクタリング改善効果の可視化）**
  リファクタリング前後の `vba-analyzer` 出力を比較し、指標の変化を定量的に表示する。
  `FOR_AI.md` の Step 6「アナライザーで今回の変更を確認する」を自動化・視覚化する機能。

  **使い方のイメージ:**
  ```bash
  # リファクタリング前にスナップショットを保存
  node test-libs/vba-analyzer.cjs src/vba/Main.bas --json > .vba-baseline.json

  # リファクタリング後に差分を表示
  node test-libs/vba-analyzer.cjs src/vba/Main.bas --diff .vba-baseline.json
  ```

  **差分表示の出力例:**
  ```
  === vba-analyzer diff: Main.bas ===

  [ScanLockedRows]
    lineCount:     37 → 22  (-15) ✅
    maxNestDepth:   5 →  3  (-2)  ✅
    flags:         DEEP_NEST, LARGE → (none) ✅

  [AccumulateLockedRowUsage]  ← 新規追加
    lineCount:     10
    flags:         (none) ✅

  [ワークスペース集計]
    totalLines:   394 → 371  (-23)
    deepNestCount:  3 →   1  (-2)  ✅
    deadCode:       2 →   2  (変化なし)
    duplicateBlocks: 2 →  0  (-2)  ✅
  ```

  **実装方針:**

  | フェーズ | 内容 | 難度 |
  |---|---|---|
  | 1. `--json` スナップショット保存 | すでに `--json` オプションあり。`.vba-baseline.json` として保存するヘルパーコマンドを追加 | ★☆☆ |
  | 2. `--diff <baseline>` モード | 現在の JSON 出力と baseline を関数名でキー照合し、数値差分・フラグ変化を計算して表示 | ★★☆ |
  | 3. 改善/悪化の色分け | 数値が改善（減少）なら ✅ 緑、悪化なら ⚠ 赤、変化なしはグレーで表示 | ★☆☆ |
  | 4. 新規追加・削除関数の検出 | baseline にない関数名 = 抽出で新規追加、baseline にあって現在ない = 削除済み | ★☆☆ |

  **実装上のポイント:**
  - baseline との照合は関数名（`name`）をキーにする。ファイル名も含めてキーを作ると複数ファイル対応になる
  - `--diff` は `--json` と組み合わせて `--diff baseline.json --json` で差分自体を JSON 出力できると AI 連携に便利
  - `.vba-baseline.json` は `.gitignore` に追加してコミット対象外にする

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
LSP 機能として VS Code に組み込む前の **プロトタイプの場** であり、ここで使い物になった機能を
順次 LSP 化していく。

### 使い方

```bash
./node_modules/.bin/esbuild test-libs/vba-analyzer.ts --bundle --outfile=test-libs/vba-analyzer.cjs --platform=node
node test-libs/vba-analyzer.cjs <file-or-dir>           # テキスト形式（人間が読む用）
node test-libs/vba-analyzer.cjs <file-or-dir> --json    # JSON 形式（プログラム/AI 連携用）
```

### 現在できること

`sample/src/vba_legacy/TaskScheduler_v1.bas` および `sample/src/vba/` での実証結果を例に。

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
| **接頭辞クラスター検出**（UDT/Enum 抽出候補）| v1: `COL_`×6, `CONFIG_`×4, `ROW_`×3 を検出 |
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

- [ ] **ロケール自動判定による出力言語切り替え**
  現在の出力（テキスト・JSON 両形式）はラベル・メッセージが日本語固定。
  AI（英語モデル）にそのまま貼り付けるときや、英語圏のチームと共有するときに障壁になる。

  **言語判定の優先順序:**
  1. 環境変数 `VBA_ANALYZER_LANG`（明示指定: `en` / `ja`）
  2. `LC_ALL` → `LC_MESSAGES` → `LANG` の順に確認（Unix/macOS）
  3. Windows: `Intl.DateTimeFormat().resolvedOptions().locale` または `process.env.LANG`
  4. いずれも取得できない場合は英語にフォールバック

  ```typescript
  function detectLang(): 'en' | 'ja' {
      const override = process.env.VBA_ANALYZER_LANG;
      if (override === 'en' || override === 'ja') return override;
      const locale = process.env.LC_ALL ?? process.env.LC_MESSAGES ?? process.env.LANG
                  ?? Intl.DateTimeFormat().resolvedOptions().locale ?? '';
      return locale.toLowerCase().startsWith('ja') ? 'ja' : 'en';
  }
  const T = detectLang() === 'ja' ? STRINGS_JA : STRINGS_EN;
  ```

  **対象範囲:**
  - `formatFileReport` / `formatWorkspaceSummary` / `formatCommentedCodeBlocks` などの人間向けテキスト出力
  - JSON 出力の `message` 等の文字列値。フィールド名はすでに英語なので変更不要

  **用途:**
  - 日本語 Windows（`LANG=ja`）では日本語、英語環境では英語が自動で出る
  - 英語出力にしたい場合は `VBA_ANALYZER_LANG=en node vba-analyzer.cjs ...` で上書き可能
  - 英語出力はトークン効率が若干よくなるケースがある（全角文字・絵文字の削減）

- [ ] **データフロー解析（Def-Use チェーン）**
  - 「この行範囲を抽出する場合、引数として渡すべき変数」を機械的に算出
  - 「この行範囲で書き換えている変数」（ByRef 引数候補）を機械的に算出
  - 抽出リファクタリング（Extract Function）の安全性を担保する基礎
  - 出力: 行範囲を指定すると `inputs[]` / `outputs[]` / `local[]` を返す API を追加

- [x] **識別子の接頭辞クラスター検出**
  - `COL_*` × 6件 / `CONFIG_*` × 4件 / `ROW_*` × 3件を検出（TaskScheduler_v1.bas での実証）
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

- [ ] **グローバル変数の抽出とスコープ解析**
  - モジュールレベル変数（`Dim` / `Public` at module scope）を一覧化
  - 各変数が参照・代入されているプロシージャ名とファイル名を列挙
  - 単一ファイル内のみで使われるか、複数ファイルにまたがるかを判定して参照数と共に出力
  - 出力: `globalVars: [{ name, scope, declFile, referencedIn: [{ file, proc, line }][], fileCount }]`
  - 用途: 引数化（§2）の候補として「多くの関数から参照されているグローバル変数」を提示

- [ ] **Type まとめ候補の検出（パラメーター共起解析）**
  - 複数の関数に同じパラメーターの組が繰り返し渡されているパターンを検出
  - またはモジュールレベル変数の同じ組が複数の関数で一緒に使われているパターンを検出
  - 出力: `udtCandidates: [{ fields: string[], usedInProcs: string[], reason: 'param-co-occurrence' | 'global-co-occurrence' }]`
  - 用途: `Type` 定義へのまとめ（§3）を具体的に提案する

- [ ] **テーブル駆動リファクタリング候補の検出（vba-analyzer への統合）**
  現在は `test-libs/table-driven-detector.ts` + `table-driven-detector-cli.ts` として独立ツールで実装・改善中。
  使い方の詳細は `test-libs/table-driven-detector-cli.ts` の `--help` を参照。
  改善が落ち着いたタイミングで `vba-analyzer` に統合し、他の検出指標（LARGE / DEEP_NEST 等）と並んで出力できるようにする。

  **統合イメージ:**
  ```
  [GetApprover]  lines=71  nest=2  ⚡ TABLE_DRIVEN_CANDIDATE(98/100)
    → department × amount の5部門・4段階をテーブル駆動化で35%削減可能
  ```

  **統合時の作業:**
  - [ ] `TableDrivenDetector` を `vba-analyzer.ts` から呼び出す
  - [ ] `--decision-table` フラグで ASCII デシジョンテーブルを出力
  - [ ] `--json` 出力に `tableDrivenCandidates[]` フィールドを追加

- [ ] **クラス抽出候補の検出（Type 使用関数のクラスタリング）**
  - 上記 `udtCandidates` で提案された `Type` を引数に取る関数が複数ある場合、それらをクラスのメソッド候補として提示
  - 出力: `classCandidates: [{ typeName, methods: [{ name, file, line }][] }]`
  - 用途: `Type` + モジュール関数からクラスへの昇格（§4）を提案する

- [ ] **`getTypeDefinitions()` を活用した型安全テスト自動生成（AI 支援）**
  - `vbaRunner.getTypeDefinitions()` の出力をプロンプトに渡し、AI に `interface` 定義とラッパー関数を生成させる
  - 生成対象: `interface <TypeName> { ... }` と `function call<ProcName>(p: <TypeName>): <ReturnType>` のペア
  - `getTypeDefinitions()` は TypeScript の型として直接使えないため（ランタイム値のため）、AI による変換ステップが必要
  - 利用イメージ:
    ```
    // 1. 出力を取得
    console.log(vbaRunner.getTypeDefinitions());
    // => { InventoryParams: { CurrentStock: 'number', SoldUnits: 'number', ... } }

    // 2. この出力を AI に渡して「型安全なラッパー関数を書いて」と依頼する
    // => interface + ラッパー関数 + Partial テストヘルパーを生成してもらう
    ```
  - 将来的には Claude Code スキル（`/generate-vba-wrappers` など）として定型化できる

- [x] **重複ブロック検出（AST N-gram マッチング）**
  - プロシージャ内・プロシージャ間で同一パターンが繰り返すステートメント列を検出
  - 変数名は `$ID` に正規化済みのため、名前が違っても構造が同じなら一致
  - スライディングウィンドウのノイズ除去（同一 proc 内のオーバーラップを先着優先で排除）
  - より長いパターンに包含される短いパターンを除去（maximal match フィルター）
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
- どこからも呼ばれていないプロシージャ（Dead code の候補。Excel ボタン・イベントハンドラー・`Application.Run` からの呼び出しは検出できないため削除前に要確認）
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

### VS Code 拡張機能（extension.ts に接続済み）
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
- [x] `package.json`（VS Code extension metadata）、`.vscodeignore`

---

## 保留中・将来課題

### npm パッケージ化と VS Code 拡張統合

> **当面はエンジン本体の改修を優先するため実施しない。**
> 構想として記録しておく。

#### 目的

現在のリポジトリ直接利用では、ユーザーのテストファイルが

```typescript
import { VBARunner } from '../../test-libs/test-runner';
```

というローカル相対パスに依存している。`vba-runner` として npm 公開することで、

```typescript
import { VBARunner } from 'vba-runner';
```

という標準的な import になり、VS Code 拡張からの利用・ユーザープロジェクトへの組み込みが自然に行える。

#### npm パッケージとして公開するもの

| エクスポート | 用途 |
|---|---|
| `VBARunner` クラス | メインの実行 API |
| `assert` | テスト用アサーション |
| `vbaTrue`, `vbaFalse`, `vbaNull`, `vbaEmpty` | VBA 型定数 |
| 型定義（`.d.ts`）| TypeScript サポート |

`Evaluator` / `Parser` / `Lexer` などのエンジン内部は原則非公開。高度なユースケース向けに別エントリで公開することは可能。

#### VS Code 拡張との統合アーキテクチャ

```
npm publish → vba-runner パッケージ
                 ↓
VS Code 拡張（vba-runner を bundle）
                 ↓ esbuild でコンパイル（ユーザーの test.ts）
                 ↓ vba-runner → ユーザーの node_modules から解決
                 ↓ child_process / worker_thread で実行
                 ↓ stdout をキャプチャ
VS Code Test API に結果を反映
```

ユーザーのテスト実行フロー（拡張機能が自動化）:
1. `.ts` テストファイルを保存
2. 拡張がワークスペースの `./node_modules/.bin/esbuild` でバンドル
3. `child_process` で `node` を起動、stdout をキャプチャ
4. `[PASS]` / `[FAIL]` を Test Explorer に反映

#### esbuild の扱い

ユーザーが `npm install --save-dev vba-runner` すると `esbuild` も devDependency として入るようにする。
拡張機能はワークスペースの `./node_modules/.bin/esbuild` を使うため、拡張自体に esbuild をバンドルする必要がない。

#### vba-analyzer の npm 化

`vba-analyzer` も同様に npm パッケージ（CLI + API）として公開する。

```bash
npx vba-analyzer <file-or-dir>          # テキスト出力
npx vba-analyzer <file-or-dir> --json   # JSON 出力（AI 連携用）
```

VS Code 拡張は `vba-analyzer` の API を import して Call Graph・Dead code・Outline 生成に使う。

#### プロセス分離（安全性）

VBA の無限ループや未捕捉例外が Extension Host プロセスを落とさないよう、
VBA 実行・TypeScript テスト実行は **`child_process` または `worker_threads`** に分離する。
現在の `Evaluator` には実行時間・メモリ上限がないため、分離は必須要件。

```
Extension Host
├── LSP / Diagnostics / 補完（軽量・同期）← 現状のまま
└── VBA実行 / テスト実行 / vba-analyzer
    └── worker_threads or child_process（ここで分離）
```

#### 残る制限・前提条件

| 項目 | 内容 |
|---|---|
| **デスクトップ限定** | `NodeFileSystem` が Node.js `fs` に依存。Web 拡張（vscode.dev）には対応しない（`MemoryFileSystem` に統一すれば将来対応可能） |
| **バージョン整合** | 拡張が bundle する `vba-runner` とユーザーの `node_modules` のバージョンが異なると挙動差が生じる可能性がある |
| **`console.log` のリダイレクト** | `VBARunner` コンストラクターに渡す print 関数を stdout 経由で OutputChannel に転送する仕組みが必要（現状の設計で対応可能） |

#### 作業順序（実施時）

- [ ] `package.json` の `exports` フィールド整備（ESM + CJS デュアルビルド）
- [ ] 公開 API の整理（`test-libs/test-runner.ts` の内容を `src/index.ts` に移動）
- [ ] `.d.ts` 型定義の出力確認
- [ ] `vba-analyzer` のエントリーポイント整備（`bin` フィールド追加）
- [ ] VS Code 拡張の `esbuild` 呼び出し実装（ワークスペースの esbuild を使う）
- [ ] VBA 実行の `worker_threads` 分離
- [ ] npm publish / VS Code Marketplace 公開

### VBA フォーマッター（`vba-formatter`）

VBA コードを AST ベースで整形する CLI ツール。`vba-analyzer` と同様に、**中核ロジックとインターフェースを分離した設計**にし、CLI・VS Code 拡張の両方から利用できるようにする。

#### アーキテクチャ方針

```
src/lsp/formatter.ts          ← 中核: FormatRule[] を受け取りTextEdit[] を返す純粋関数
    ↑ import
test-libs/vba-formatter.ts    ← CLI エントリーポイント（vba-analyzer と同様）
src/extension.ts              ← DocumentFormattingEditProvider から formatter.ts を呼ぶ
```

- **`formatter.ts`（中核）**: `format(source: string, options: FormatterOptions): TextEdit[]` を公開。AST と元ソースから編集箇所のリストを返す。VS Code の `WorkspaceEdit` や CLI の文字列置換の両方に対応できる形式。
- **CLI（`vba-formatter.ts`）**: `--check`（差分表示のみ）/ `--write`（上書き）モードを持つ。`vba-analyzer` と同様に `./node_modules/.bin/esbuild` でバンドルして実行。
- **VS Code 拡張**: `registerDocumentFormattingEditProvider` から `formatter.ts` を呼ぶ。ルールの ON/OFF は VS Code 設定（`vba.formatter.*`）で制御。

#### 実装予定のフォーマットルール

| ルール | デフォルト | 内容 |
|---|---|---|
| `indentation` | ON | `If`/`For`/`Sub` 内を4スペース（または指定幅）でインデント |
| `keywordCase` | ON | キーワードを正規大文字化（`dim` → `Dim`、`sub` → `Sub`） |
| `operatorSpacing` | ON | 演算子前後にスペース（`x=1` → `x = 1`） |
| `thenOnSameLine` | ON | `If x` の後の `Then` を同一行に |
| `lineWrap` | ON | 指定幅（デフォルト 120 文字）を超える行を `_` 継続で折り返す |
| `argumentAlignment` | ON | 折り返した引数リストを開き括弧の直後の位置に揃える（下記参照） |
| `mergeDimAndAssign` | OFF | `Dim` 宣言の直後に同一変数への代入がある場合、`: ` で1行にまとめる（下記参照） |
| `moveDimToFirstUse` | OFF | `Dim` 宣言を初使用直前に移動（VBA はプロシージャスコープのため意味論的に安全。大規模ファイルでは diff が大きくなるためデフォルト OFF） |

#### `lineWrap` / `argumentAlignment` の変換例

```vb
' 変換前（長い関数呼び出し）
result = SomeLongFunctionName(firstArgument, secondArgument, thirdArgument, fourthArgument)

' 変換後（開き括弧の直後に揃える）
result = SomeLongFunctionName(firstArgument, _
                              secondArgument, _
                              thirdArgument, _
                              fourthArgument)
```

- 折り返し位置はトークン境界（引数の `,` の後）で行う
- 継続行の頭は開き括弧 `(` の次の文字位置にスペースで揃える
- 既存の `_` 継続行はいったん展開してから再整形する（べき等性のため）

#### `mergeDimAndAssign` の変換例

```vb
' 変換前
Dim total As Long
total = CalcTotal(a, b)

' 変換後（: で1行に）
Dim total As Long: total = CalcTotal(a, b)
```

- 対象: `Dim` の直後の文（間に他の文がない場合のみ）が同一変数への代入
- 除外: `Set` 代入（オブジェクト参照）、`Dim x As New Foo`、配列宣言

#### `moveDimToFirstUse` の除外ケース

- `Dim x As New Foo`（Auto-Instantiation — オブジェクト生成タイミングに意図がある）
- 配列宣言（`Dim arr(10) As Long`）
- `Dim` と初使用の間に `On Error` がある場合

#### Diagnostics: `Dim` 型指定の罠（警告）

フォーマッターとは別に、LSP Diagnostics（または vba-analyzer）として以下の警告を追加する。

**`Dim x, y, z As Long` — 先頭変数が Variant になる罠**

VBA では `Dim x, y, z As Long` と書いた場合、型指定 `As Long` は **最後の `z` にのみ適用**される。`x` と `y` は `Variant` になる。他の言語（C, TypeScript 等）から来た開発者が誤りやすい。

```vb
Dim x, y, z As Long   ' ⚠️ x, y は Variant（Long ではない）
```

```
警告: 'x', 'y' は型指定がなく Variant になります。
     各変数に型を明示してください: Dim x As Long, y As Long, z As Long
```

- 検出条件: `VariableDeclaration` で宣言子が2つ以上あり、かつ途中の宣言子に `As` 型指定がない
- severity: Warning（意図的な Variant 利用の可能性があるため Error にしない）
- 実装場所: `getDiagnostics()` に追加（`CodeLensProvider.getDeadCodeWarnings()` と同様の統合方式）

#### 作業順序（実施時）

- [x] `src/lsp/formatter.ts` の中核実装（`indentation` + `keywordCase` から着手）
- [x] `test-libs/vba-formatter.ts` CLI 実装（`--check` / `--write` / stdout）
- [x] VS Code 拡張への統合（`DocumentFormattingEditProvider` 登録）
- [x] フォーマッター・警告のテスト（`tests/lsp/lsp-formatter.test.ts`）
- [ ] `lineWrap` + `argumentAlignment` ルール実装
- [ ] `mergeDimAndAssign` ルール実装
- [ ] `moveDimToFirstUse` ルール実装（Def-Use 解析または参照検索で初使用行を特定）
- [x] `Dim x, y, z As Long` 型指定警告を `getDiagnostics()` に追加（VBA001 として vba-lint.ts に実装済み）
- [ ] GoTo ラベル（`Label:` 形式）はインデントしない（現在は本文と同じ深さになっている）
- [ ] **型名の大文字化**（`keywordCase` ルールの拡張）
  現在のレキサーは `Long`/`Integer`/`Single`/`Double`/`String`/`Boolean`/`Byte`/`Currency`/`Date`/`Object`/`Variant` を識別子（`Identifier` トークン）として扱うため、`dim x as long` の `long` が `Long` に正規化されない。
  VBA標準のスタイルではこれらも Pascal 化が期待される。
  **対応方針:** `formatter.ts` の `KEYWORD_CANONICAL` に型名エントリを追加し、識別子トークンの値（小文字化）が型名と一致する場合に正規化する。トークンは単語単位なので `stringVar` のような変数名と混同しない。
  対象型名: `Integer`, `Long`, `Single`, `Double`, `String`, `Boolean`, `Byte`, `Currency`, `Date`, `Object`, `Variant`, `LongLong`, `LongPtr`

### パッケージング（Marketplace 公開）
- [ ] VS Code Marketplace への公開（アイコン整備、`.vsix` ビルド検証）
- [ ] Web Extension 化（evaluator が Node.js `path` に依存しており、先にブラウザ対応が必要）

### Web UI（デモサイト）の改善
- [ ] `Dir` 関数の完全実装（ディレクトリ列挙）
- [ ] `Kill` ワイルドカード対応（例: `Kill "*.txt"`）
- [ ] コールスタックの出力・`Erl` 関数サポート

---

**最終ビジョン**:
VBA開発者が VBE を開かずに VS Code だけで LSP 補完を受けながらコードを書き、参照・リネームで安全にリファクタリングし、テストを即座に実行し、ステップデバッグで挙動を確認できる「モダンなVBA開発環境」を実現する。

さらに、数万行のレガシーVBAに対して、コードベース全体をAIに効率よく把握させ、
「どこから手をつけるか」「どう分割するか」をAIと対話しながら進められる
**AI支援リファクタリングの起点** となることを目指す。
