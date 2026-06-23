# EVALUATION.md — 評価パイプラインの詳細

VBA コードが `callProcedure` で実行されるまでに、3 パス＋実行のフェーズを経る。

```
ソーステキスト
    │  ① パース（字句解析 → 構文解析）
    ▼
AST (Program)
    │  ② Pass 1: evaluateModule — シンボルテーブル構築
    ▼
env に手続き・変数が登録された状態
+ pendingArrayDecls（配列境界付き Dim の退避リスト）
+ pendingTopLevel（モジュールレベル実行文の退避リスト）
    │  ③ Pass 2: resolveIdentifiers — 識別子解決・実行準備
    ▼
すべての定数・配列が確定し、モジュールレベル実行文も済んだ状態
resolveIdentifiersDone = true
    │  ④ Pass 3: precheckProc — AST 静的チェック（コンパイルエラー）
    ▼
静的チェック通過済み
    │  ⑤ 実行: executeStatements
    ▼
プロシージャ本体の実行
```

---

## ① パース

`Lexer` がソーステキストをトークン列に変換し、`Parser` が AST を構築する。
この段階では env への登録も式の評価も行わない。

---

## ② Pass 1 — `evaluateModule(program)` : シンボルテーブル構築

**目的**: 全モジュールの「名前と構造」を env に登録する。式の評価は最小限にとどめる。

マルチモジュールの場合、全モジュールを `evaluateModule` でロードしてから Pass 2 を呼ぶ。
これにより「モジュール A の定数がモジュール B の定数を参照する」といったクロスモジュール参照に対応できる。

### 各文の処理

| 文の種別 | Pass 1 での扱い | 理由 |
|---|---|---|
| `ProcedureDeclaration` (Sub/Function/Property) | env に即時登録 | 本体は実行しない |
| `ConstDeclaration` | **スキップ** | 他の Const を参照する場合に依存順が不明なため |
| `VariableDeclaration`（配列境界なし） | env に即時登録 | 初期値は型デフォルト値のみで Const 参照なし |
| `VariableDeclaration`（配列境界あり） | **`pendingArrayDecls` に退避** | 境界式が Const を参照する可能性があるため |
| `TypeDeclaration` / `EnumDeclaration` / `ClassDeclaration` | 即時登録 | — |
| Option 系 / DefType / Attribute / Declare / Event / Implements | 即時評価 | コンテキスト設定のため Pass 1 で必要 |
| **モジュールレベル実行文**（代入・呼び出し・制御文など） | **`pendingTopLevel` に退避** | Const・配列 Dim が確定していないため |

### バッチモードとインタラクティブモード

`resolveIdentifiersDone` フラグで動作が切り替わる。

- **`false`（バッチモード）**: `evaluateModule` 呼び出し中。上表の通り退避を行う。
- **`true`（インタラクティブモード）**: Pass 2 完了後。以降の `evaluateModule` 呼び出し（`evalExpression` のフォールバックや追加モジュールロード）はすべて即時実行する。

### 関連フィールド

```ts
private resolveIdentifiersDone = false;
private pendingArrayDecls: Array<{ stmt: VariableDeclaration; moduleName: string }> = [];
private pendingTopLevel:   Array<{ moduleName: string; stmts: Statement[] }> = [];
```

---

## ③ Pass 2 — `resolveIdentifiers(modules)` : 識別子解決・実行準備

全モジュールのロード完了後に **1 回だけ** 呼ぶ。以下の順序で処理する。

### 2-1. 定数のトポロジカルソートと評価

全モジュールの `ConstDeclaration` を収集し、依存グラフを構築してトポロジカルソートする。
循環参照はここでエラーになる。クロスモジュール参照も正しく解決される。

```
Const A = B + 1   ← B に依存
Const B = 10      ← 依存なし

ソート結果: B → A の順に評価
```

他モジュールの `Private Const` を参照しようとした場合もエラー（VBA 仕様: Private は同一モジュール内のみ）。

### 2-2. 配列境界付き Dim の評価（`pendingArrayDecls`）

Const がすべて env に確定した後、`pendingArrayDecls` を順に評価する。
これにより `Dim a(0 To W - 1)` のような境界式が正しく解決される。

```vb
Const W As Integer = 3
Dim a(0 To W - 1) As Integer  ' Pass 1 では退避、Pass 2 で W=3 確定後に評価 → UBound=2
```

### 2-3. 重複プロシージャ名チェック

同一モジュール内で同名の Sub/Function が宣言されていないかを検証する（コンパイルエラー）。

### 2-4. Option Explicit 静的チェック（キャッシュ構築）

全モジュール名が確定した状態で `checkOptionExplicit` を実行し、未宣言変数の使用箇所を `optionExplicitViolations` にキャッシュする。  
この段階では throw せず、実際のエラー発生は Pass 3（`precheckProc`）に委ねる。  
これにより `ev.set()` による実行時注入後にチェックを行える（テスト・モック注入との整合性）。

### 2-5. `resolveIdentifiersDone = true` に設定

以降の `evaluateModule` 呼び出しをインタラクティブモードに切り替える。

### 2-6. モジュールレベル実行文の実行（`pendingTopLevel`）

Const・配列 Dim がすべて確定した状態でモジュールレベル実行文を順次実行する。
ロード順（`evaluateModule` が呼ばれた順序）を保持する。

```vb
Const INIT = 42
Dim g As Integer
g = INIT   ' ← pendingTopLevel に退避 → Pass 2 末尾で実行 → g = 42
```

---

## ④ Pass 3 — `precheckProc` : AST 静的チェック（コンパイルエラー）

`precheckProc` はプロシージャ本体を実行する前に AST を静的に走査するフェーズで、ここで検出されるエラーはコンパイルエラーとして報告される。チェックの例として、Option Explicit 違反、Sub を関数とみなして戻り値を取得しようとしていないか、未定義プロシージャの呼び出し、Dim の配列サイズに非定数式が使われていないか、GoTo/GoSub の未定義ラベル、呼び出し引数の不一致などがある。

## ⑤ 実行 — `executeStatements`

静的チェックがすべてパスした後、プロシージャ本体（AST の `body` 文リスト）を `executeStatements` で実行する。

---

## Class 宣言の評価（インライン `Class...End Class` と `.cls` の関係）

### Pass 1 では「クラス名の登録」しかしない

`ClassDeclaration` は `isModuleLevelDeclaration`（`src/engine/evaluator.ts:2151`）で
モジュールレベル宣言と判定されるため、Pass 1 中に即時評価される。だが
`evaluateClassDeclaration`（`evaluator.ts:3373`）の実体はこれだけである。

```ts
private evaluateClassDeclaration(stmt: ClassDeclaration) {
    this.registerClass(stmt.name, stmt);
}
```

`registerClass`（`evaluator.ts:3369`）は AST ノード（`stmt` そのもの、つまり
`fields`/`procedures`/`body` を含む全体）を `classDefinitions` マップにクラス名（小文字）
キーで保存するだけ。**フィールドの既定値設定・Property Get/Let/Set の対応付け・
クラス内 Const の評価・Sub/Function 本体の検証は、Pass 1 では一切行われない。**
これらはすべて `New ClassName` によるインスタンス化時まで遅延される。

### インスタンス化時（`New`）にクラス本体を評価する

`instantiateClass`（`evaluator.ts:3377`）が `classDefinitions` から AST を取り出し、
`createInstanceFromDef`（`evaluator.ts:3392`）が実質的に「クラス本体に対する
Pass 1 + Pass 2」をこの時点でまとめて行う。

1. `instanceEnv = new Environment(this.env)` でインスタンス専用の環境を作成
2. `classDef.fields` を走査し、型に応じた既定値（`""`/`0`/`vbaEmpty`）を `instanceEnv` に設定。`WithEvents` フィールドはここでマーキング
3. `__classDef__`/`__instanceEnv__`/`__events__` を持つインスタンスオブジェクトを構築
4. `classDef.body` から `EventDeclaration` を走査し、イベントスロットを初期化
5. `Me` を `instanceEnv` にセット
6. **「クラス本体の Pass 2」**: `classDef.body` 内の `ConstDeclaration` を `evaluateConstValue` で評価し `instanceEnv` に定数として設定（コード上のコメントは `B-1: Evaluate class-level Const declarations into instanceEnv`）
7. `classDef.procedures` に `Class_Initialize` があれば呼び出す

モジュールレベルの `pendingArrayDecls`（§Pass 2 の 2-2）に相当する仕組みはクラスの
フィールドには存在しない。フィールドの既定値設定（手順 2）は配列境界の有無にかかわらず
スカラーの既定値しか設定しないため、配列境界付きフィールド（`Dim a(0 To N) As Integer`
のようなクラスフィールド）の境界式評価は別経路になっている（本書では未調査）。
Sub/Function/Property 本体そのものは、メソッド呼び出し時に `classDef.procedures` から
都度引き当てられるだけで、インスタンス化時に事前評価されることはない。

### インライン `Class...End Class` と `.cls` ファイルは完全に同じ経路

どちらも同じ `parseClassBody(className, untilEndClass)`（`src/engine/parser.ts:2064`）
が返す `ClassDeclaration` ノードを使う。

| 形式 | 呼び出し元 | `untilEndClass` | 結果 |
|---|---|---|---|
| インライン `Class Foo ... End Class`（`__mocks__.bas` 等） | 通常の文パース中に `KeywordClass` を検出（`parser.ts:1657`） | `true`（`End Class` まで読む） | `Program.body` の中の 1 要素としての `ClassDeclaration` |
| `.cls` ファイル | `new Parser(tokens, { parseAsClass: moduleName })`（`VBARunner`: `test-libs/test-runner.ts:296`、VS Code 拡張: `src/extension.ts:506`） | `false`（EOF まで読む） | `parse()`（`parser.ts:1276`）が `{ type: 'Program', body: [classDecl], diagnostics }` を返す。つまり `Program.body` がその `ClassDeclaration` 1 個だけの配列 |

`ClassDeclaration` 自体のノード形状（`fields`/`procedures`/`body`）はどちらも同一であり、
`parseAsClass` は「`Program.body` の包み方」と「終端条件（`End Class` か EOF か）」だけを
変える。`evaluateModule` 以降は `program.body` を順に見るだけなので、`ClassDeclaration` の
出自（インラインか `.cls` か）を区別するコードはどこにもない。**つまりインライン Class と
`.cls` の評価は完全に同じ経路であり、上記の「Pass 1 ではクラス名のみ登録・インスタンス化時に
本体を評価」という挙動はどちらにも同様に当てはまる。**

なお、`ClassDeclaration` は標準 VBA の `.bas` ファイル構文には存在しない（`.cls` は
別ファイル種別）。インラインの `Class...End Class` は vba-runner 独自拡張であり、
§「プロシージャの後に書いた宣言・実行文のチェック」の位置チェックでも対象外にしている
（`evaluator.ts` の `resolveIdentifiers` 内、`stmt.type === 'ClassDeclaration'` は
スキップ）。

---

## 呼び出し元ごとのフロー

### `evalVBASingle(code)` — 単一モジュールテスト

```
parse → evaluateModule(ast)         ← Pass 1
      → resolveIdentifiers([{ast}]) ← Pass 2（pendingTopLevel も実行）
      → return Evaluator            ← callProcedure は呼び出し側が行う
```

### `evalVBAModules(modules)` — マルチモジュールテスト

```
for each module:
    parse → evaluateModule(ast)         ← Pass 1（全モジュール分）
→ resolveIdentifiers([...all asts...])  ← Pass 2（1回のみ）
→ return Evaluator
```

### `VBARunner` (test-libs) — ファイルベーステスト

```
for each .bas/.cls file:
    parse → evaluateModule(ast)         ← Pass 1
→ 初回 run()/eval() 呼び出し時に _ensureResolved() が
  resolveIdentifiers(...) を遅延実行      ← Pass 2（遅延）
```
> Pass 2 が遅延している間に `ev.set()` / `ev.setConstant()` で定数や変数を注入できる。

### `evalExpression(expr)` — インタラクティブ評価

Pass 2 完了後（`resolveIdentifiersDone = true`）に呼ばれる想定。
`evaluateModule` を内部で呼ぶが、インタラクティブモードのため退避せず即時実行する。

---

## トップレベル文の評価（vba-runner 拡張）

### 背景

標準の VBA 仕様では、モジュールのトップレベルに書けるのは宣言文（`Dim`・`Const`・`Sub`・`Function`・`Type`・`Enum` 等）のみであり、代入や呼び出しといった実行文は Sub/Function の本体にしか書けない。

vba-runner はこれを拡張しており、モジュールトップレベルに実行文を書くことができる。主な用途は `evalVBASingle` を使った簡易テストスクリプトやプレイグラウンドの REPL である。

```vb
' ← これが「トップレベル実行文」
Const W = 3
Dim a(0 To W - 1) As Integer
a(0) = 99   ' ← 標準 VBA では Sub の外に書けないが vba-runner は許容する
```

### Pass 1 での扱い（退避）

トップレベル実行文をその場で実行すると、Const や配列 Dim がまだ確定していない状態になる（Pass 1 時点では Const はスキップ済み）。そのため Pass 1 ではモジュールごとに `pendingTopLevel` リストへ退避するだけにとどめる。

```
Pass 1: g = INIT を pendingTopLevel に退避（INIT はまだ未解決）
Pass 2: Const INIT = 42 を評価 → env に確定
        pendingTopLevel を実行 → g = 42 ✓
```

実行順序はモジュールのロード順（`evaluateModule` の呼び出し順）で保たれる。複数モジュール間でトップレベル文がある場合、ロード順に実行される。

### 呼び出し元ごとの違い

| 呼び出し元 | トップレベル実行文の扱い |
|---|---|
| `evalVBASingle` / `evalVBAModules` | `pendingTopLevel` に退避 → Pass 2 末尾で実行 |
| Playground (`App.tsx`) | AST をフィルタリングして宣言文のみを `evaluateModule` に渡す（実行文は除外） |
| `evalExpression` | Pass 2 完了後の呼び出しのため即時実行（インタラクティブモード） |

#### Playground のフィルタリング

Playground は「ソースコード」ペインと「テストスクリプト」ペインを別々に扱う。ソースコード側の VBA は Sub/Function などの宣言のみを `evaluateModule` に渡し、トップレベル実行文は意図的に除外する。

```ts
// App.tsx
body: ast.body.filter(stmt =>
    stmt.type === 'ProcedureDeclaration' ||
    stmt.type === 'TypeDeclaration'      ||
    stmt.type === 'VariableDeclaration'  ||
    stmt.type === 'ConstDeclaration'
)
```

テストスクリプト（TypeScript）側は `vbaRunner.eval(exprString)` を通じて `evalExpression` を呼び出し、1 行ずつインタラクティブに実行する。

#### `evalExpression` のインタラクティブ実行

`resolveIdentifiersDone = true`（Pass 2 完了後）の状態で `evaluateModule` が呼ばれると、バッチモードではなくインタラクティブモードとして動作し、実行文は退避せず即時実行される。

ただし `evalExpression` は文字列を式（`Expression`）として解析しようとする。`g = 100` はVBA の `=` が代入と比較の両方に使われるため、式としてパースされた場合は比較式（`False`）として評価される。文として明示的に実行したい場合は `vbaRunner.run('Foo', [])` か、別途 Sub にラップする必要がある。

### 制約

- モジュールレベル実行文は `pendingTopLevel` に退避されるため、**Pass 1 中**（他のモジュールロード中）には実行されない。モジュール間の初期化順序依存を避けるため、これは意図的な挙動である。
- `VBARunner`（ファイルベース）は Pass 2 を遅延する（`_ensureResolved`）。そのため `run()`/`eval()` が呼ばれるまでトップレベル実行文も実行されない。この期間に `ev.set()` でモック変数を注入できる。

---

## プロシージャの後に書いた宣言・実行文のチェック（`allowTopLevelStatements`）

標準 VBA では、モジュール内で最初のプロシージャ（`Sub`/`Function`/`Property`）が始まった後は
`End Sub`/`End Function`/`End Property` の後に書けるのはコメントのみである。
`Dim`/`Const`/`Type`/`Enum` などの宣言文や、代入・`For`/`Next` などの実行文を書くと
`コンパイル エラー: End Sub、End Function または End Property 以降には、コメントのみが記述できます。`
という VBA のコンパイルエラーになる（`tests/vba/CompileError.bas` の `module_level_*_after_procedure` 系ケース）。

vba-runner は §「トップレベル文の評価」で説明した拡張（モジュールレベル実行文を許容する）を、
**プロシージャの前後どちらに書いても**有効にしている。つまり「先頭にない `Dim`」も
トップレベル実行文と同じ扱いで許容される。`evalVBASingle` を使った簡易テストスクリプトでは
`Function F() ... End Function` の後に `Dim v As Variant` を書く、といった書き方が広く使われて
いるため、これは意図的なデフォルト動作である。

`Evaluator` の `allowTopLevelStatements` オプション（デフォルト `true`）でこの拡張の有効/無効を
切り替えられる。`false` にすると、最初のプロシージャより後に他の文（宣言文・実行文どちらも）が
現れた時点で標準 VBA 相当のコンパイルエラーになる。チェック自体は `resolveIdentifiers`（Pass 2）
で行われる（`src/engine/evaluator.ts` の `resolveIdentifiers` 内、重複プロシージャ名チェックの直後）。

### 呼び出し元ごとの `allowTopLevelStatements` の既定値

「VBA の再現・評価」が目的の経路（`VBARunner`・VS Code 拡張・Playground のテストペイン）は
すべて `false`（標準 VBA 相当の厳密チェック）に統一している。一方、`evalVBASingle`/`evalVBAModules`
（テストスクリプト用の簡易実行）と Playground の「ソースコードを実行する」機能は、モジュール
トップレベルの実行文をそのまま動かすことが目的のため `true`（拡張あり）のままにしている。

| 呼び出し元 | 既定値 | 切り替え可否 | 備考 |
|---|---|---|---|
| `evalVBASingle` / `evalVBAModules`（`test-libs/test-runner.ts`） | `true` | `options.allowTopLevelStatements` で呼び出しごとに上書き可能 | `tests/spec/` の簡易テストスクリプト用。`Function F() ... End Function` の後に `Dim`/代入を書く書き方を許容する。`CompileError.bas` の strict 系ケースは `EVAL_OPTIONS: { allowTopLevelStatements: false }` で明示的に上書きしている |
| `VBARunner`（ファイルベース、`sample/tests/ts/` 用、`test-libs/test-runner.ts` のコンストラクター） | `false` 固定 | 不可（ハードコード） | 実 `.bas`/`.cls` ファイルの VBA 再現・評価が目的のため、標準 VBA 仕様に違反するコードは検出する |
| VS Code 拡張の実行コマンド（`vba-runner.runProcedure`、`src/extension.ts`） | `false` 固定 | 不可 | `ev.reEvaluateModuleConstsAll(asts)`（`resolveIdentifiers` の旧名エイリアス）経由で Pass 2 に到達する |
| VS Code 拡張のテスト実行（`src/lsp/test-runner.ts`）・デバッグ実行（`src/lsp/debug-worker.ts`） | `false` 固定 | 不可 | いずれも実ファイルを評価して実行するため `VBARunner`/実行コマンドと同じ方針 |
| LSP 診断（エディター上の赤波線、`src/lsp/server.ts` の `getDiagnostics`） | 該当なし | — | 診断は `Lexer`/`Parser.parse()` のみで行われ、`Evaluator`/Pass 2 にまったく到達しない。つまりこのチェック（`TYPE: prerun`）はエディター上の波線としては**絶対に表示されない**。波線になるのは構文エラー（`TYPE: parse`）のみ |
| Playground の「ソースコードを実行する」機能（`handleRun`、`src/App.tsx`） | `true` 固定 | 不可 | これはスクリプトランナー機能であり、モジュールトップレベルの `Sub` 呼び出し（デフォルトスニペットの末尾の `MainLoop` 等）を実行することが目的のため、意図的に拡張を有効にしている |
| Playground のテストペイン（`createBrowserTestRunner` が作る `vbaRunner.run`/`vbaRunner.eval`、`src/App.tsx`） | `false` 固定 | 不可 | `VBARunner` と同じ「VBA の再現・評価」用途のため strict 化した。ソースコード側 AST を宣言文（`ProcedureDeclaration`/`TypeDeclaration`/`VariableDeclaration`/`ConstDeclaration`）のみにフィルターしてから `evaluateModule` に渡す点は変更していないが、`Dim`/`Const` はこのフィルターを通過するため、プロシージャの後に書くと `allowTopLevelStatements: false` により正しくコンパイルエラーになる |

実用上のまとめ:

- VBA コードの実行・テスト・デバッグを担う経路（`VBARunner`・VS Code 拡張・Playground テストペイン）は
  すべて標準 VBA 相当の厳密チェックが既定で有効。プロシージャの後に `Dim`/`Const`/実行文を書くと
  これらの経路ではコンパイルエラーになる。
- `evalVBASingle`/`evalVBAModules` と Playground の「ソースコードを実行する」機能だけは、
  簡易テストスクリプト・REPL 的な用途のため拡張を有効のままにしている。

---

## なぜ「3 パス」か

静的解析パスは `evaluateModule`（Pass 1）・`resolveIdentifiers`（Pass 2）・`precheckProc`（Pass 3）の 3 つで、その後にプロシージャ本体の実行（`executeStatements`）が続く。

| # | 名前 | メソッド | 主な処理 |
|---|---|---|---|
| 1 | Parse | `Lexer` + `Parser` | AST 構築（評価なし） |
| 2 | Pass 1 | `evaluateModule` | シンボルテーブル構築・モジュールレベル実行文を pendingTopLevel に退避（→ [トップレベル文の評価](#トップレベル文の評価vba-runner-拡張)） |
| 3 | Pass 2 | `resolveIdentifiers` | 定数解決・配列初期化・静的チェック・実行文実行 |
| 4 | Pass 3 | `precheckProc` | 各 Proc 実行直前に AST 静的チェック（コンパイルエラー） |
| 5 | 実行 | `executeStatements` | プロシージャ本体の実行 |
