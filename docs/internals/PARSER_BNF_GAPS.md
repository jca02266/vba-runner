# Parser BNF Gap Analysis — §5.4 Statements

BNF.md（MS-VBAL §5.4）とパーサー実装を照合した結果。
調査日: 2026-07-14（再調査: 2026-07-15 / ステータス更新: 2026-07-15）

---

## §5.4.2 Control Statements

| 節 | 状態 | 問題・VBA例 |
|---|---|---|
| §5.4.2.1 Call | ⚠️ | `Call .Method()` (parens あり) は ✅。`Call .Method` / `Call .Method arg`（parens なし）は ImplicitWithObjectExpression が CallExpression でないためエラー |
| §5.4.2.2 While | ✅ | — |
| §5.4.2.3 For | ✅ | `Next i, j` 複数変数 → 2026-07-15 修正済み (`pendingNextVars` キュー方式) |
| §5.4.2.4 For Each | ✅ | §5.4.2.3 と同根。同一修正で対応 |
| §5.4.2.5 Exit For | ✅ | — |
| §5.4.2.6 Do | ✅ | — |
| §5.4.2.7 Exit Do | ✅ | — |
| §5.4.2.8 If (multiline) | ✅ | — |
| §5.4.2.9 Single-line If | ⚠️ | `If x > 0 Then 100`（行番号ジャンプ）が GoTo として機能しない |
| §5.4.2.10 Select Case | ⚠️ | `Case Is ><5` `Case Is =>10` `Case Is =<5`（代替比較演算子）未対応 |
| §5.4.2.11 Stop | ✅ | — |
| §5.4.2.12 GoTo | ⚠️ | `Go To MyLabel`（2語形式）未対応 → `Go` が識別子扱いになりパースエラー |
| §5.4.2.13 On…GoTo | ✅ | — |
| §5.4.2.14 GoSub | ⚠️ | `Go Sub MyLabel`（2語形式）未対応（§5.4.2.12 と同根） |
| §5.4.2.15 Return | ✅ | — |
| §5.4.2.16 On…GoSub | ✅ | — |
| §5.4.2.17 Exit Sub | ✅ | — |
| §5.4.2.18 Exit Function | ✅ | — |
| §5.4.2.19 Exit Property | ✅ | — |
| §5.4.2.20 RaiseEvent | ✅ | — |
| §5.4.2.21 With | ✅ | — |
| §5.4.2.22 End | ✅ | — |
| §5.4.2.23 Assert | ⚠️ | `Debug.Assert` が `AssertStatement` ノードでなく `CallStatement` になる（実行は正しい、LSP用途のみ問題） |

---

## §5.4.3 Data Manipulation Statements

| 節 | 状態 | 問題・VBA例 |
|---|---|---|
| §5.4.3.1 Dim/Static | ⚠️ | `Dim Shared x As Integer`（Shared キーワード）未対応（VBA6方言、実用頻度低） |
| §5.4.3.2 Const | ✅ | `Const A = 1, B = 2` 複数定数 → 2026-07-15 修正済み |
| §5.4.3.3 ReDim | ✅ | `ReDim arr1(5), arr2(10)` 複数変数 → 2026-07-15 修正済み |
| §5.4.3.3 ReDim | ✅ | `ReDim obj.Arr(5)` メンバーアクセス対象 → 2026-07-15 修正済み |
| §5.4.3.3 ReDim | ✅ | `ReDim .Items(n)` With式対象 → 2026-07-15 修正済み |
| §5.4.3.4 Erase | ✅ | `Erase arr1, arr2` 複数要素 → 2026-07-15 修正済み |
| §5.4.3.5 LSet | ✅ | `LSet s = expr` 対応済み |
| §5.4.3.6 Mid | ✅ | — |
| §5.4.3.7 RSet | ✅ | — |
| §5.4.3.8 Let | ⚠️ | `Let x = 5` は動作するが `Let` が AST に記録されない（実害なし） |
| §5.4.3.9 Set | ✅ | — |

---

## §5.4.4 Error Handling Statements

| 節 | 状態 | 問題・VBA例 |
|---|---|---|
| §5.4.4.1 Error | ✅ | — |
| §5.4.4.2 Resume | ✅ | — |
| §5.4.4.3 On Error | ✅ | `On Error GoTo -1` → 2026-07-15 修正済み |
| §5.4.4.3 On Error | ✅ | `On Error GoTo 0` / `On Error Resume Next` — 正常 |

---

## §5.4.5 File Statements

| 節 | 状態 | 問題・VBA例 |
|---|---|---|
| §5.4.5.1 Open | ✅ | `Open "f.txt" As #1`（`For mode` 省略）→ 2026-07-15 修正済み |
| §5.4.5.1 Open | ⚠️ | `Open "f.txt" For Random Access Read Shared As #1`（`Shared` スタンドアロン）→ `Lock` キーワード後にのみ `Shared` を読むため未対応 |
| §5.4.5.1 Open | ✅ | `Len = N` 句（Bug 26-3 修正済み） |
| §5.4.5.1.1 File Numbers | ✅ | `Seek 1, 100` / `Print 1, "x"` 等（`#` なし）→ 2026-07-15 修正済み（8文すべて対応） |
| §5.4.5.2 Close/Reset | ✅ | — |
| §5.4.5.3 Seek | ✅ | `#` なし対応（§5.4.5.1.1 と同根、同一修正） |
| §5.4.5.4 Lock | ⚠️ | `Lock #1, To 100`（start 省略の `To end` 形式）→ `start = parseExpression()` を先読みするため未対応 |
| §5.4.5.5 Unlock | ⚠️ | Lock と同根 |
| §5.4.5.6 Line Input | ⚠️ | `variable-name` が Identifier のみ（`Line Input #1, arr(i)` 不可） |
| §5.4.5.7 Width | ✅ | `#` なし対応（§5.4.5.1.1 と同根） |
| §5.4.5.8 Print | ✅ | `#` なし対応（§5.4.5.1.1 と同根） |
| §5.4.5.9 Write | ✅ | `#` なし対応。output-list 内の Spc/Tab は未処理（式として評価される可能性あり） |
| §5.4.5.10 Input | ✅ | `#` なし対応（§5.4.5.1.1 と同根） |
| §5.4.5.11 Put | ✅ | `#` なし対応（§5.4.5.1.1 と同根） |
| §5.4.5.12 Get | ✅ | `#` なし対応（§5.4.5.1.1 と同根） |

---

## 優先度別サマリー

### 🔴 高（実用コードで踏む可能性が高い）

なし（2026-07-15 全件修正済み）

### 🟠 中（実用コードで踏む可能性あり）

なし（2026-07-15 全件修正済み）

### 🟡 低（レガシーまたは稀な書き方）

| 問題 | 場所 | VBA例 |
|---|---|---|
| `Go To` / `Go Sub` 2語形式 | §5.4.2.12/14 | 現代 VBA では1語が主流 |
| `Case Is ><5` 代替比較演算子 | §5.4.2.10 | VBA IDE でも警告が出る書き方 |
| `If x Then 100` 行番号ジャンプ | §5.4.2.9 | レガシー BASIC スタイル |
| `Open ... Shared` スタンドアロン | §5.4.5.1 | ネットワークファイル操作 |
| Lock/Unlock `To end`（start省略） | §5.4.5.4/5 | `Lock #1, To 100` |
| `Line Input #1, arr(i)` | §5.4.5.6 | variable が配列要素 |
| `Dim Shared` | §5.4.3.1 | VBA6 方言 |
| `Call .Method arg`（parens なし With形式） | §5.4.2.1 | With ブロック内 Call |

### ⬜ 無視可（動作は正しい）

| 問題 | 場所 | 補足 |
|---|---|---|
| `Let x = 5` が AST に `Let` 未記録 | §5.4.3.8 | 実行に影響なし |
| `Debug.Assert` が `CallStatement` ノード | §5.4.2.23 | LSP の専用処理が必要な場合のみ問題 |

---

# Parser BNF Gap Analysis — §3.4 + §5.2–§5.3 Declarations

BNF.md（MS-VBAL §3.4 + §5.2–§5.3）とパーサー実装を照合した結果。
調査日: 2026-07-15（ステータス更新: 2026-07-15）

---

## §3.4.1 Conditional Compilation Const Directive

| 節 | 状態 | 問題・VBA例 |
|---|---|---|
| §3.4.1 #Const | ✅ | `#Const DEBUG = True` → `src/engine/preprocessor.ts` でソーステキストを Lexer 前処理（初回監査時の調査漏れ） |

---

## §3.4.2 Conditional Compilation If Directives

| 節 | 状態 | 問題・VBA例 |
|---|---|---|
| §3.4.2 #If | ✅ | `#If DEBUG Then ... #End If` → §3.4.1 と同根。`preprocessor.ts` で対応済み |

---

## §5.2 Module Declaration Section Structure

| 節 | 状態 | 問題・VBA例 |
|---|---|---|
| §5.2 モジュール宣言構造 | ⚠️ | 宣言セクション・コードセクションの明示的分離なし。パーサーは全体をフラットに処理。機能上の問題はないが仕様の厳密さを欠く |

---

## §5.2.1 Option Directives

| 節 | 状態 | 問題・VBA例 |
|---|---|---|
| §5.2.1.1 Option Compare | ✅ | — |
| §5.2.1.2 Option Base | ✅ | — |
| §5.2.1.3 Option Explicit | ✅ | — |
| §5.2.1.4 Option Private Module | ✅ | — |

---

## §5.2.2 Implicit Definition Directives (Def*)

| 節 | 状態 | 問題・VBA例 |
|---|---|---|
| §5.2.2 DefInt / DefStr 等 | ✅ | — |

---

## §5.2.3 Module Declarations

| 節 | 状態 | 問題・VBA例 |
|---|---|---|
| §5.2.3.1 モジュール変数宣言 | ✅ | `Global x As Integer` → 2026-07-15 修正済み（`KeywordGlobal` 追加 + `CONTEXTUAL_KW_STRUCTURAL` に追加） |
| §5.2.3.1.1 変数宣言 | ✅ | — |
| §5.2.3.1.2 WithEvents | ✅ | — |
| §5.2.3.1.3 配列次元と境界 | ✅ | — |
| §5.2.3.1.4 変数型宣言 | ✅ | — |
| §5.2.3.2 Const 宣言 | ✅ | `Const A = 1, B = 2` → 2026-07-15 修正済み（バッチ1 §5.4.3.2 と同根） |
| §5.2.3.3 UDT 宣言 (Type) | ⚠️ | メンバー名として `End` が使用不可。ループ終端が `KeywordEnd` チェックのため `End As Long` で早期終了 |
| §5.2.3.4 Enum 宣言 | ✅ | `Global Enum MyEnum` → 2026-07-15 修正済み（`Global` 対応により `parseEnumDeclaration` に到達） |
| §5.2.3.5 外部プロシージャ宣言 | ✅ | `Declare [PtrSafe] Sub/Function ... Lib ... [Alias ...]` 対応済み |

---

## §5.2.4.1.2 Implements / Event

| 節 | 状態 | 問題・VBA例 |
|---|---|---|
| §5.2.4.1.2 Implements | ⚠️ | `Implements SomeLib.IInterface`（ドット修飾名）→ `parseImplementsDirective` は1トークンしか読まず `.IInterface` がパースエラー |
| §5.2.4.1.2 Event | ✅ | — |

---

## §5.3 Module Code Section Structure

| 節 | 状態 | 問題・VBA例 |
|---|---|---|
| §5.3 プロシージャ宣言全般 | ✅ | Sub / Function / Property Get/Let/Set、Friend スコープ含め対応済み |

---

## 優先度別サマリー（バッチ2 更新版）

### 🔴 高（実用コードで踏む可能性が高い）

なし（2026-07-15 全件修正済み）

### 🟠 中（実用コードで踏む可能性あり）

なし（2026-07-15 全件修正済み）

### 🟡 低（レガシーまたは稀な書き方）

| 問題 | 場所 | VBA例 |
|---|---|---|
| `Implements SomeLib.IInterface` ドット修飾未対応 | §5.2.4.1.2 | `Implements ADODB.ICommand` |
| UDT メンバー名 `End` 使用不可 | §5.2.3.3 | `Type T: End As Long: End Type` |

---

# Parser BNF Gap Analysis — §5.3 + §5.4.1 Procedure Declarations + Statement Blocks

BNF.md（MS-VBAL §5.3.1–§5.3.1.10 + §5.4–§5.4.1）とパーサー実装を照合した結果。
調査日: 2026-07-15

---

## §5.3.1 Procedure Declarations

| 節 | 状態 | 問題・VBA例 |
|---|---|---|
| §5.3.1 手続き宣言（全体） | ✅ | Sub / Function / Property Get/Let/Set 対応済み |
| §5.3.1.1 Procedure Scope | ✅ | `Global Sub Foo()` → 2026-07-15 修正済み（`Global` = `Public` スコープ） |
| §5.3.1.2 Static Procedures | ✅ | 先頭 `Static Sub` と末尾 `Sub Foo() Static` 両方対応 |
| §5.3.1.3 Procedure Names | ⚠️ | `Function Foo$()` 等 TYPED-NAME サフィックス付き宣言名 → `Foo$` のまま登録。`Foo()` と呼び出すと名前不一致になる可能性あり |
| §5.3.1.4 Function Type Declarations | ✅ | `As Type` / `As Type()` 対応。ドット修飾1段階まで |
| end-label | ✅ | プロシージャ末尾のラベルは body の LabelStatement として処理 |

---

## §5.3.1.5 Parameter Lists

| 節 | 状態 | 問題・VBA例 |
|---|---|---|
| procedure-parameters | ✅ | — |
| property-parameters | ⚠️ | `Property Let Foo(x As Integer, v As String)` の value-param 強制チェックなし（意味論的制約、パース自体は可） |
| positional / optional params | ✅ | `Optional`, `ByVal`, `ByRef`, `ParamArray` 対応 |
| optional-prefix 順序 | ⚠️ | `ByVal Optional x As Integer`（ByVal が先）→ `parseParameter` は Optional を先に消費するため ByVal→Optional 順のとき `Optional` をパラメーター名として読もうとしパースエラー |
| typed-name-param-dcl | ⚠️ | `Function Foo(x% As Integer)` — x% の型サフィックスがパラメーター名に含まれる。§5.3.1.3 と同根 |
| parameter-type "Any" | ✅ | `As Any` は文字列として記録される |

---

## §5.4.1 Statement Blocks

| 節 | 状態 | 問題・VBA例 |
|---|---|---|
| §5.4.1 statement-block | ✅ | — |
| §5.4.1.1 識別子ラベル | ✅ | `MyLabel:` 対応 |
| §5.4.1.1 行番号ラベル | ✅ | `100:` または `100`（コロンなし）対応 |
| §5.4.1.2 Rem statement | ✅ | レキサーレベルで消費されパーサーには届かない |
| attribute-statement | ✅ | `Attribute VB_Name = "..."` 対応 |
| `:` EOS 区切り | ✅ | `isAtTerminator()` が OperatorColon を終端として扱う |

---

## 優先度別サマリー（バッチ3）

### 🔴 高（実用コードで踏む可能性が高い）

なし

### 🟠 中（実用コードで踏む可能性あり）

なし

### 🟡 低（レガシーまたは稀な書き方）

| 問題 | 場所 | VBA例 |
|---|---|---|
| `ByVal Optional x` パラメーター順序未対応 | §5.3.1.5 | `Sub Foo(ByVal Optional x As Integer)` |
| TYPED-NAME サフィックス付き宣言名 | §5.3.1.3 | `Function Foo$()` を `Foo()` で呼ぶと不一致 |

---

# Parser BNF Gap Analysis — §5.6.5–§5.6.9 式・演算子

BNF.md（MS-VBAL §5.6.5–§5.6.9）とパーサー実装を照合した結果。
調査日: 2026-07-15（ステータス更新: 2026-07-15）

---

## §5.6.5 Literal Expressions

| 節 | 状態 | 問題・VBA例 |
|---|---|---|
| §5.6.5 整数・浮動小数点・日付・文字列 | ✅ | INTEGER / FLOAT / DATE / STRING 対応済み |
| §5.6.5 literal-identifier (True/False/Empty/Null/Nothing) | ✅ | Identifier として処理、評価器で解決 |
| §5.6.5 `D` 指数表記 | ⚠️ | `1.5D10`（VBA Double 指数）→ レキサーは `e`/`E` のみ処理、`D`/`d` 未対応。`1.5` + 識別子 `D10` に分割される（§3.x レキサー問題） |

---

## §5.6.6 Parenthesized Expressions

| 節 | 状態 | 問題・VBA例 |
|---|---|---|
| §5.6.6 `(expression)` | ✅ | — |

---

## §5.6.7 TypeOf…Is Expressions

| 節 | 状態 | 問題・VBA例 |
|---|---|---|
| §5.6.7 TypeOf…Is | ✅ | `TypeOf x Is ADODB.Recordset` ドット修飾型名 → 2026-07-15 修正済み |

---

## §5.6.8 New Expressions

| 節 | 状態 | 問題・VBA例 |
|---|---|---|
| §5.6.8 `New TypeName` | ✅ | `New ADODB.Recordset` 等1段階ドット修飾まで対応 |

---

## §5.6.9 Operator Expressions

| 節 | 状態 | 問題・VBA例 |
|---|---|---|
| §5.6.9.3 算術演算子 (+, -, *, /, \, Mod, ^, &) | ✅ | — |
| §5.6.9.5 標準比較演算子 (=, <>, <, >, <=, >=) | ✅ | — |
| §5.6.9.5 代替比較演算子 (`><`, `=<`, `=>`) | ⚠️ | バッチ1 §5.4.2.10 と同根。レキサーが `><` を `>` + `<` 別々に出力するためパース不可 |
| §5.6.9.6 Like | ✅ | — |
| §5.6.9.7 Is | ✅ | — |
| §5.6.9.8 論理演算子 (Not, And, Or, Xor, Eqv, Imp) | ✅ | — |
| 比較演算子の優先度 | ⚠️ | `a = b < c` → パーサーは `a = (b < c)` と解析（`=` を `<` より低優先度扱い）。VBA仕様では全比較演算子は同一優先度で左結合 → `(a = b) < c` が正しい（連鎖比較は実用上極めて稀） |

---

## 優先度別サマリー（バッチ4）

### 🔴 高（実用コードで踏む可能性が高い）

なし

### 🟠 中（実用コードで踏む可能性あり）

なし

### 🟡 低（レガシーまたは稀な書き方）

| 問題 | 場所 | VBA例 |
|---|---|---|
| `D` 指数表記未対応（レキサー問題） | §5.6.5 | `1.5D10` |
| 比較演算子の優先度誤り（連鎖比較） | §5.6.9.5 | `a = b < c` の解析順 |
| 代替比較演算子 `><`/`=<`/`=>` | §5.6.9.5 | バッチ1既報と同根 |

---

# Parser BNF Gap Analysis — §5.6.11–§5.6.16 式（メンバーアクセス・With・AddressOf 等）

BNF.md（MS-VBAL §5.6.11–§5.6.16）とパーサー実装を照合した結果。
調査日: 2026-07-15

---

## §5.6.11 Instance Expressions

| 節 | 状態 | 問題・VBA例 |
|---|---|---|
| §5.6.11 `Me` | ✅ | `Me` → Identifier として処理 |

---

## §5.6.12 Member Access Expressions

| 節 | 状態 | 問題・VBA例 |
|---|---|---|
| §5.6.12 `obj.Property` | ✅ | postfix ループで MemberExpression に変換 |
| §5.6.12 行継続後 `.` | ✅ | レキサーが行継続 `_` を消費するため問題なし |
| §5.6.12 `unrestricted-name`（予約語プロパティ） | ✅ | `obj.Print`、`ws.Get` 等、COMPAT_KW_EXPR で許可済み |

---

## §5.6.13 Index Expressions / Argument Lists

| 節 | 状態 | 問題・VBA例 |
|---|---|---|
| §5.6.13 `f(args)` index-expression | ✅ | CallExpression として統一処理 |
| §5.6.13.1 空の位置引数 `f(, x)` | ✅ | MissingArgument ノードで対応 |
| §5.6.13.1 名前付き引数 `name := value` | ✅ | NamedArgument ノードで対応 |
| §5.6.13.1 `argument-expression = ["byval"] expression` | ⚠️ | `Foo(ByVal x)` → `ByVal` が呼び出し引数では未処理でパースエラー（実用頻度は極めて低い） |
| §5.6.13.1 `AddressOf` in argument | ✅ | `Foo(AddressOf Bar)` 対応済み |

---

## §5.6.14 Dictionary Access Expressions

| 節 | 状態 | 問題・VBA例 |
|---|---|---|
| §5.6.14 `obj!Key` | ✅ | postfix ループで DictionaryAccessExpression に変換 |

---

## §5.6.15 With Expressions

| 節 | 状態 | 問題・VBA例 |
|---|---|---|
| §5.6.15 `.Property`（With メンバーアクセス） | ✅ | ImplicitWithObjectExpression として処理 |
| §5.6.15 `!Key`（With 辞書アクセス） | ⚠️ | `With dict: !Key = v: End With` → parsePrimary は行頭 `!` を With-dictionary-access として処理しない。`!` が単体ではエラー |

---

## §5.6.16 その他の式カテゴリ

| 節 | 状態 | 問題・VBA例 |
|---|---|---|
| §5.6.16.1 constant-expression | ✅ | expression の別名 |
| §5.6.16.2 cc-expression | ✅ | 条件付きコンパイル式。`preprocessor.ts` で対応済み |
| §5.6.16.3–5 boolean/integer/variable-expression | ✅ | expression / l-expression の別名 |
| §5.6.16.7 type-expression（多段階ドット） | ⚠️ | `Foo As A.B.C` → 2段階以上は未対応（バッチ2〜4の共通問題） |
| §5.6.16.8 AddressOf（1段階ドット） | ✅ | `AddressOf Module.Proc` 対応済み |
| §5.6.16.8 AddressOf（多段階ドット） | ⚠️ | `AddressOf NS.Module.Proc` → 1段階のみ対応（実用上は稀） |

---

## 優先度別サマリー（バッチ5）

### 🔴 高（実用コードで踏む可能性が高い）

なし

### 🟠 中（実用コードで踏む可能性あり）

なし

### 🟡 低（レガシーまたは稀な書き方）

| 問題 | 場所 | VBA例 |
|---|---|---|
| `With dict: !Key = v` With 辞書アクセス行頭 `!` 未対応 | §5.6.15 | With ブロック内 `!Key` |
| `Foo(ByVal x)` 呼び出し側 ByVal 未対応 | §5.6.13.1 | 古い COM スタイルの呼び出し |
| type-expression の多段階ドット | §5.6.16.7 | `Dim x As A.B.C`（バッチ2〜4と同根） |

---

# 全体サマリー（2026-07-15 更新）

## 🔴 高優先度の残件

**なし**

## 🟠 中優先度の残件

**なし**

## 🟡 低優先度の残件（修正は任意）

| 問題 | 場所 | VBA例 |
|---|---|---|
| `Go To` / `Go Sub` 2語形式 | §5.4.2.12/14 | レガシー |
| `Case Is ><5` 代替比較演算子 | §5.4.2.10/§5.6.9.5 | VBA IDE でも警告 |
| `If x Then 100` 行番号ジャンプ | §5.4.2.9 | レガシー BASIC |
| `Open ... Shared` スタンドアロン | §5.4.5.1 | ネットワーク専用 |
| `Lock/Unlock #1, To 100` start省略 | §5.4.5.4/5 | 稀な書き方 |
| `Line Input #1, arr(i)` 配列要素 | §5.4.5.6 | 稀な書き方 |
| `Dim Shared x` VBA6方言 | §5.4.3.1 | レガシー |
| `Call .Method arg` With形式 parens なし | §5.4.2.1 | 稀な書き方 |
| `Implements SomeLib.IInterface` ドット修飾 | §5.2.4.1.2 | COM ライブラリ実装時 |
| UDT メンバー名 `End` 使用不可 | §5.2.3.3 | 稀な命名 |
| `ByVal Optional x` パラメーター順序 | §5.3.1.5 | 稀な書き方 |
| TYPED-NAME サフィックス付き宣言名 `Foo$()` | §5.3.1.3 | レガシー |
| `D` 指数表記（レキサー問題） | §5.6.5 | 科学技術計算 |
| 比較演算子の優先度誤り（連鎖比較） | §5.6.9.5 | 極めて稀 |
| `With dict: !Key = v` With辞書行頭 `!` | §5.6.15 | 稀な書き方 |
| `Foo(ByVal x)` 呼び出し側 ByVal | §5.6.13.1 | 古い COM スタイル |
| type-expression 多段階ドット | §5.6.16.7 | `Dim x As A.B.C` |
| `AddressOf NS.Module.Proc` 多段階 | §5.6.16.8 | 稀な書き方 |
