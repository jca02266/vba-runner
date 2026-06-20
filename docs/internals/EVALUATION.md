# EVALUATION.md — 評価パイプラインの詳細

VBA コードが `callProcedure` で実行されるまでに、4 つのフェーズを経る。

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
    │  ④ Call time: callProcedure → precheckProc → 実行
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
この段階では throw せず、実際のエラー発生は Call time（`precheckProc`）に委ねる。  
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

## ④ Call time — `callProcedure` → `precheckProc` → 実行

プロシージャを呼び出す際に `precheckProc` で実行時チェックを行う。

### `precheckProc` の処理

1. **Option Explicit 違反チェック**: Pass 2 でキャッシュした `optionExplicitViolations` と、現在の env・`defaultBindingObject` を照合する。実行時注入済みの変数は違反から除外される。
2. **Sub が値として使われていないかチェック** (`checkSubAsValueInProc`)
3. **未定義プロシージャ呼び出しチェック** (`checkUndefinedCallsInProc`)

これらがすべてパスした後、プロシージャ本体（AST の `body` 文リスト）を `executeStatements` で実行する。

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

## なぜ「2.5 パス」か

厳密には Pass は 2 つ（`evaluateModule` と `resolveIdentifiers`）だが、
Pass 2 の末尾でモジュールレベル実行文を動かすため、
論理的には「2 パス + 実行」の 2.5 パス構成に近い。

Call time の `precheckProc` を独立したパスと数えると 3 パスとも言える。

| # | 名前 | メソッド | 主な処理 |
|---|---|---|---|
| 1 | Parse | `Lexer` + `Parser` | AST 構築（評価なし） |
| 2 | Pass 1 | `evaluateModule` | シンボルテーブル構築・実行文退避 |
| 3 | Pass 2 | `resolveIdentifiers` | 定数解決・配列初期化・静的チェック・実行文実行 |
| 4 | Call time | `precheckProc` + `execProcBody` | 実行時検証・プロシージャ実行 |
