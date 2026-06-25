# AMBIGUOUS_SYNTAX.md — VBA の構文的曖昧性と VBA Runner の対処

VBA の文法には、見た目だけでは複数の解釈が可能な構文がいくつかある。実 VBA では
「文として書かれた場合」と「式として書かれた場合」で解釈が一意に決まるため、
通常の `.bas`/`.cls` ファイルを書く・読む側はこの曖昧性を意識する必要がない。

しかし VBA Runner の `evalExpression()`（`VBARunner.eval()` の実体）は、
**入力全体を「単一の式」として一発評価できないか試す高速パス**を持っている
（`eval("1 + 2 + 3")` が即座に `6` を返せるようにするための設計）。この高速パスは
「式として解析できるか」だけを見るため、**本来は文として解釈すべき入力を
誤って式として確定させてしまう**ことがある。本書はその曖昧性の種類と、
VBA Runner が実際にどう判定しているかをまとめる。

バグ修正・`eval()` 周辺の機能追加時に参照すること。

## 大前提: 文として解析する限り曖昧性は存在しない

`.bas`/`.cls` ファイルをロードして `run()` で呼ぶ通常の経路（`evaluateModule` →
`evaluateStatement` / `executeStatements`）では、ソースは常に**文（statement）の並び**
として `Parser.parseStatement()` 系で解析される。文の文法は曖昧性なく一意に決まり、
後述する全パターンは「常に同じ解釈」になる。

曖昧性が表面化するのは、`evalExpression()` が「これは式かもしれない」と
`parseExpressionPublic()` を先に試す、**そのときだけ**である。

```
                          ┌─ parseExpressionPublic() で全体を消費できる？
eval(exprString) ─────────┤    Yes → 式として評価して値を返す（高速パス）
                          └─ No  → 文の並びとして parse → executeStatements()
                                   （戻り値は常に undefined）
```

`run()`/`.bas` ファイル経由は常に下のルート（文として解析）だけを通るため、
以下の曖昧性は **`eval()` 専用の問題** であることに注意。

---

## 1. ラベル vs 文区切りの `:` — `a : b`

**規則**: 識別子（または contextual keyword）の直後のトークンが `:` の場合、
**空白の有無に関係なく常にラベル宣言**として解釈される。`a : b` ・ `a:b` ・
`a :b` はすべて同一のトークン列（`Identifier(a)`, `OperatorColon`, `Identifier(b)`）
になるため、レキサーの時点で空白の情報自体が失われており、パーサーが空白の有無で
挙動を変える余地はそもそもない。

```vb
a : b     ' => LabelStatement "a" + CallStatement "b"
a: b      ' => 同上（空白の有無は無関係）
a:b       ' => 同上
```

数値リテラルは `:` が**なくても**常にラベルとして解釈される:

```vb
10        ' => LabelStatement "10"（コロンなしでも）
10: GoTo 10
```

一方、識別子が括弧・ドット・代入等を伴う場合はラベルとして認識されない
（通常の文区切りの `:` になる）:

```vb
Foo() : b   ' => CallStatement "Foo()" + CallStatement "b"（ラベルではない）
a.b : c     ' => CallStatement + CallStatement
a(1) : c    ' => CallStatement + CallStatement
x = 1 : y   ' => AssignmentStatement + CallStatement（後述の#2と組み合わさるケース）
```

**実装**: `parser.ts` の `parseIdentifierOrCallStatement()`。

```ts
// Label check: "Identifier:" or contextual-keyword ":"
if ((token.type === TokenType.Identifier || Parser.CONTEXTUAL_KW.has(token.type)) &&
        this.pos + 1 < this.tokens.length &&
        this.tokens[this.pos + 1].type === TokenType.OperatorColon) {
    // ラベルとして確定
}

// Line number label: "42" or "42:"
if (token.type === TokenType.Number) {
    // コロンの有無に関わらずラベルとして確定（コロンは this.match() で任意消費）
}
```

「次のトークンが `:` か」だけを見る単純な先読みであり、`parsePrimary()` で
式を構築する**前**に判定が行われる。そのため `Foo()` のように一度でも括弧や
ドットを消費した時点でこの分岐には到達せず、ラベル化されない。

`eval()` への影響は小さい（`eval('a : b')` のように裸の識別子をラベルにしたい
状況自体が稀）が、エラーメッセージの行数推定やデバッグ時に「なぜ `a` が
変数として読まれないのか」と混乱した場合はこの規則を疑うこと。

---

## 2. `=` — 代入 vs 等価比較

**規則**:
- **文として解析される場合**: 常に代入。`parseIdentifierOrCallStatement()` は
  `parsePrimary()`（identifier/call/member access まで。**二項演算子は消費しない**）
  で左辺を構築した直後に `=` を見るので、構造的に比較演算子として解釈される
  余地がない。
  ```ts
  const expr = this.parsePrimary(); // foo, foo(), foo.bar, arr(0) など
  if (this.match(TokenType.OperatorEquals)) {
      return { type: 'AssignmentStatement', left: expr, right: this.parseExpression() };
  }
  ```
- **式として解析される場合**（`If x = 10 Then` の条件式、関数引数など）:
  常に等価比較。`parseExpressionPublic()` は二項演算子としての `=` を
  含む完全な式文法をそのまま解釈する。

**`eval()` での問題**: `evalExpression()` の高速パスは `parseExpressionPublic()`
を**先に**試すため、`x = 10` という入力は「`x` と `10` を比較する
`BinaryExpression`」として全体をきれいに消費できてしまう。これが文だと
判定されないと、代入が一切実行されず比較結果の真偽値だけが返る
（2026-06-25 に発見・修正したバグ。詳細は `TODO_SPEC.md` の
「テストランナー (`VBARunner`) の改善」セクション参照）。

**対処**: `evalExpression()` で、フルに消費された式のトップレベル演算子が
`=` の `BinaryExpression` の場合は高速パスを採用せず、文として解析する
フォールバックに委ねる（`isStatementAmbiguous` の判定。`src/engine/evaluator.ts`
の `evalExpression()` 内）。これにより `eval('x = 10')` も `eval('arr(1) = "a"')`
も正しく代入として実行されるようになった。

---

## 3. `+`/`-` で始まる識別子主導の式 — 算術 vs 暗黙の Call 文

VBA の `<implicit-call-statement>` 文法は、識別子（や `obj.Member`・`arr(i)` の
ような呼び出し可能な形）で始まる裸の文を「`Call` キーワードなし・括弧なしの
手続き呼び出し」として解釈する。引数は空白区切りで、各引数は
`parseCallArgument()` → `parseExpression()` で**完全な式として**パースされる。
そのため `+1` や `-1` のような単項演算子始まりの式は「1個目の引数」として
正当に成立してしまう。

```vb
x - 1          ' => CallStatement: x を引数 -1 で呼ぶ
x + 1 + 2      ' => CallStatement: x を引数 (+1 + 2) で呼ぶ（深いネストも同様）
Foo(1) + 1     ' => CallStatement: Foo(1) を引数 +1 で「呼ぶ」（括弧付き呼び出しでも同様）
x.Bar + 1      ' => CallStatement: x.Bar を引数 +1 で呼ぶ
```

一方、比較・連結・論理演算子（`<`・`&`・`And` など）はそもそも
`parseExpression()` の**先頭トークンとして開始できない**ため、引数として
成立せず Parse error になる。つまりこれらの演算子には曖昧性自体が存在しない:

```vb
x < 10     ' => Parse error（"<" は式の先頭になれない。文としても成立しない）
x & "a"    ' => Parse error（同上）
x And y    ' => Parse error（同上）
```

括弧で明示的に囲むと曖昧性が解消され、常に式として解釈される
（括弧で囲んだ時点で「裸の識別子主導の文」という形ではなくなるため）:

```vb
(x) + 1    ' => 文としては Parse error。式としてのみ成立する → 常に算術加算
```

**`eval()` での問題**: 上記2.と同根。`x + 1`（`x` が変数）は高速パスで
算術式として評価されてしまい `6` を返すが、これは VBA の文法上ありえない
解釈である。さらに悪いことに、`x` が必須引数を持つ Function/Sub を指す場合、
`evaluateExpression()` のバラ識別子評価は「自動呼び出しは引数0個の場合のみ」
という規則によりこの呼び出しをスキップし、`x` を未宣言変数同様の `Empty`(0)
として読んでしまうため `0 + 1 = 1` のような完全に誤った値を返す（呼び出しは
一度も発生しない）。`eval('... : x + 1')` のように複数文の最後に置いた場合は
正しく Call 文として解釈されていたため、**単独 `eval()` と複数文中とで
結果が食い違う**という形でこのバグが発覚した（2026-06-25 修正）。

**対処**: `evaluator.ts` の `isCallableLeftmostLeaf()` ヘルパーが、`+`/`-` の
連鎖を再帰的に左へ辿り、最終的な葉が `Identifier`/`CallExpression`/
`MemberExpression`（= 呼び出し可能な形）であれば「statement として解析すべき」
と判定する。括弧で囲まれた式（`ParenthesizedExpression`）は辿る対象から
除外され、`(x) + 1` のような明示的な曖昧性解消は維持される。

```ts
private isCallableLeftmostLeaf(expr: Expression): boolean {
    let node: Expression = expr;
    while (true) {
        switch (node.type) {
            case 'Identifier':
            case 'CallExpression':
            case 'MemberExpression':
                return true;
            case 'BinaryExpression': {
                const op = (node as BinaryExpression).operator;
                if (op !== '+' && op !== '-') return false;
                node = (node as BinaryExpression).left;
                continue;
            }
            default:
                return false;
        }
    }
}
```

---

## 4. なぜ `eval()` は複数文入力の末尾の値を返さないのか

`eval('Dim x : x = 10 : x')` は `undefined` を返す（`x` の値 `10` は返らない）。
これはバグではなく、**実 VBA にそもそも対応する構文がない**ことの反映である。

実 VBA の statement 文法には「文の並びの最後の式を評価して返す」という機能が
存在しない。裸の式 `x` だけの行は、`x` が呼び出し可能でなければ文として成立せず
（`Foo` のような0引数手続きなら暗黙の Call 文として成立する。詳細は
[FUNCTION_CALL.md](FUNCTION_CALL.md) 参照）、`eval()` が「最後の文の値を返す」
ことを保証しようとすると、文法上意味を持たない概念を発明することになる。

`evalExpression()` の文として解析するフォールバックは、単に
`executeStatements(program.body, 0); return undefined;` で終わる。値が欲しい
場合は:

- `eval()` を分けて呼ぶ（`eval('Dim x : x = 10'); const v = eval('x');`）
- `Debug.Print`/`onPrint` で出力する
- 単一の式だけを渡す（高速パスが効き、値が返る）

のいずれかを使う。README 5節（`quiet`/`onPrint`）・REFERENCE.md にこの制限の
記載がある。

---

## 5. 関連する既知の罠（修正済みだが構造を理解する上で参考になる）

### `On Error Resume Next:` の直後の `:` で後続文が消える（2026-06-25 修正済み）

`parseOnErrorStatement()` の "Resume Next" 読み取りループが、`:`（文区切り）
ではなく `Newline`/`EOF` でしか止まらない実装になっていたため、
`On Error Resume Next: Debug.Print "x"` のように同一行に `:` で後続文を
書くと、`:` 以降の全トークンが label 文字列として丸ごと飲み込まれ、
AST から消えていた。これは曖昧性ではなく単純なバグだったが、「`:` を
文区切りとして正しく認識できているか」という同種の確認漏れとして
発見された。修正は既存の `isAtTerminator()`（`Newline`/`EOF`/`OperatorColon`
を終端とする）を使うよう変更。

### `eval()` 呼び出しごとの状態リセット

`evalExpression()` は呼び出しごとに `currentSourceModule` と
`errorHandlerLabel`/`errorHandlingMode`/`isInErrorHandler`/`lastErrorIndex`
を退避・リセットし、**各 `eval()` 呼び出しを独立したトップレベル呼び出し
フレームとして扱う**。これにより:

- ファイルロード後に残留した `currentSourceModule` の影響で、`eval()` 内の
  module-level `Dim` の書き込み先と読み込み元がズレる問題（2026-06-25 修正）
- 直前の `eval('On Error Resume Next')` の状態が後続の無関係な `eval()` 呼び出しに
  漏れ残る問題

の両方を防いでいる。これは実 VBA のプロシージャ呼び出しが `On Error` 状態を
呼び出し元から独立してリセットする（手続きごとにスコープを持つ）のと同じ
原則であり、`eval()` を「都度新しい手続き呼び出しに入るのと同等」として
扱っている。

---

## まとめ表

| 構文 | 文として解析した場合 | 式として解析した場合 | 曖昧性が存在するか |
|---|---|---|---|
| `a : b`（識別子 + コロン） | 常にラベル | （該当なし） | なし（空白は無関係） |
| `x = 10` | 常に代入 | 常に等価比較 | **あり**（`eval()` 高速パスで誤判定しうる） |
| `x + 1` / `x - 1`（識別子主導） | 常に暗黙 Call 文 | 常に算術演算 | **あり**（`eval()` 高速パスで誤判定しうる） |
| `x < 10` / `x & "a"` / `x And y` | Parse error（成立しない） | 常に比較・連結・論理演算 | なし（文として成立しないため） |
| `(x) + 1`（括弧で囲む） | Parse error（成立しない） | 常に算術演算 | なし（括弧で曖昧性解消済み） |
