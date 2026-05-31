# 名前解決：仕様と実装

## MS-VBAL §5.6.10 の名前空間階層

シンプルネーム式（`xlUp`、`i`、`MyFunc` など）は以下の **6 階層**を順に検索して解決される（Static Semantics）。
最初にマッチした階層が採用される。どの階層にもマッチしない場合、`explicit-mode` ならエラー、`implicit-mode` ならその手続き内にローカル変数を暗黙宣言する。

| 優先度 | 階層 | 含まれるもの |
|--------|------|-------------|
| 1 | **手続き名前空間** | ローカル変数・パラメーター・定数（`Dim`/`Const`/`Static` で宣言） |
| 2 | **モジュール名前空間** | モジュールレベル変数・定数・関数（同一モジュール内） |
| 3 | **プロジェクト名前空間** | プロジェクト自身・参照プロジェクト・同プロジェクト内モジュール名 |
| 4 | **同プロジェクト内の他モジュール** | 他の `.bas`/`.cls` で `Public` 宣言された変数・定数・関数 |
| 5 | **参照プロジェクト名前空間** | 参照プロジェクトのモジュール名 |
| 6 | **参照プロジェクトのモジュール内** | 参照プロジェクト（Excel 型ライブラリ等）の `Public` 定数・変数・関数 |

Excel VBA で `xlUp` が使えるのは Excel 型ライブラリが **tier 6** に存在するため。

---

## 当エンジンの実装対応状況

| 階層 | 対応 | 実装 |
|------|------|------|
| 1 手続き名前空間 | ✅ | `callProcedure` が作成するローカル `Environment` |
| 2 モジュール名前空間 | ✅ | `evaluateModule()` 時にモジュールレベル宣言を global `Environment` へ登録 |
| 3 プロジェクト名前空間 | ⚠️ | モジュール名によるアクセス（`Module.Proc`）は部分対応 |
| 4 同プロジェクト内の他モジュール | ✅ | 複数 `.bas` を `evaluateModule()` すると global env に蓄積される（変数・関数・`Public Const` が対象） |
| 5 参照プロジェクト名前空間 | ❌ | 未実装 |
| 6 参照プロジェクトのモジュール内 | ❌ | 型ライブラリなし。`xlUp` 等は未定義 |

tier 6 の定数（`xlUp` 等）はテスト時に `runner.set()` または別モジュールで定義して提供する。

---

## モジュールレベル定数の評価

### 二段階ロード

VBA ではプロジェクト内の全モジュールがコンパイル時に一括解決されるため、他モジュールの `Public Const` を参照する定数はロード順によらず正しく解決される（tier 4）。

当エンジンはこれを二段階ロードで実現する：

```
Pass 1: 全モジュールの evaluateModule() を実行
  → 手続き・変数を env に登録する
  → ConstDeclaration はスキップ（右辺を評価しない）
  → Option Explicit 静的解析を実行（knownModuleNames なし：保守的モード）

Pass 2: resolveIdentifiers() を呼び出す
  → 全モジュールの定数をグローバルトポロジカルソートして正しい順序で評価
  → クロスモジュール Public Const 参照はソート順で自然に解決
  → 未定義名を参照すると "Constant expression required" エラー
  → Option Explicit 静的解析を再実行（knownModuleNames あり：精密モード）
```

**VBARunner では `resolveIdentifiers()` を初回 `run()`/`eval()` 呼び出し時まで遅延する。**
これにより `set()` による外部定数の注入をコンストラクター後に行える。

```javascript
const runner = new VBARunner('./DataProcessor.bas');
runner.set('xlUp', -4162);   // ← resolveIdentifiers() の前に注入できる
runner.run('LastRow', [ws]); // ← ここで初めて resolveIdentifiers() を実行
```

### Const 右辺で参照できるもの

Const の右辺に書ける識別子は以下の通り：

| 参照先 | 可否 | 説明 |
|--------|------|------|
| リテラル（数値・文字列・`True`/`False` 等） | ✅ | 常に可 |
| 組み込み VBA 定数（`vbCrLf`、`vbOK` 等） | ✅ | 起動時に env に登録済み |
| 外部注入値（`runner.set()` / `ev.set()` で Pass 2 前に注入） | ✅ | env に登録済み |
| **同一モジュール**内の Const | ✅ | グローバルトポロジカルソートで正しい順序に解決 |
| **他モジュール**の `Public Const` | ✅ | グローバルトポロジカルソートで解決（VBA 仕様通り） |
| **他モジュール**の `Private Const` | ❌ | `Constant expression required` エラー（VBA 仕様通り） |
| どこにも定義されていない名前 | ❌ | `Constant expression required` エラー |

### 循環参照

| ケース | エラー |
|--------|--------|
| 同一モジュール内の循環（`Const A = B + 1; Const B = A + 1`） | `Circular reference in constant declarations: a → b → a` |
| クロスモジュールの循環（`Module1.A → Module2.B → Module1.A`） | `Circular reference in constant declarations: ...` |

---

## Option Explicit チェック

### 仕様

§5.6.10 より：名前解決で全階層にマッチしない場合、`explicit-mode`（`Option Explicit` あり）なら「simple name expression is invalid」。

### 実装

チェックは三段階で行う：

1. **静的解析 Pass 1**（`evaluateModule()` 時）: `checkOptionExplicit(ast)` が AST を走査し、未宣言識別子を使うプロシージャ名と未宣言名を `optionExplicitViolations: Map<procName, Set<undeclaredNames>>` に記録する。この時点では他モジュール名が未知なので、コール式の bare identifier オブジェクト（`obj.Method()` の `obj` 部分）はすべてスキップする（保守的モード）。

2. **静的解析 Pass 2**（`resolveIdentifiers()` 末尾）: `checkOptionExplicit(ast, knownModuleNames)` を全モジュールに対して再実行する。第 2 引数 `knownModuleNames` に全モジュール名の集合を渡すことで、コール式 bare identifier が **本物のモジュール名かどうかを判定**できるようになる。モジュール名でない bare identifier は未宣言変数として検出し `optionExplicitViolations` にマージする（精密モード）。

3. **呼び出し時判定**（`callProcedure` 時）: 記録された未宣言名を `env.hasVariable(name)` で再確認する。名前が env に存在すれば tier 1〜4 で解決済みとみなしてエラーを出さない。

Pass 2 完了後、定義されていない名前は env に存在しないため `hasVariable` が false を返し、Option Explicit チェックが正しく機能する。

```javascript
// xlUp が別モジュールで定義されている場合
const runner = new VBARunner('./directory/');
runner.run('LastRow', [ws]);   // Pass 2 後、env に xlup あり → OK

// runner.set() で注入する場合
const runner = new VBARunner('./DataProcessor.bas');
runner.set('xlup', -4162);
runner.run('LastRow', [ws]);   // env に xlup あり → OK

// どこにも定義しない場合
const runner = new VBARunner('./DataProcessor.bas');
runner.run('LastRow', [ws]);   // env に xlup なし → Option Explicit エラー
```

### checkOptionExplicit のシグネチャ

```typescript
checkOptionExplicit(program: Program, knownModuleNames?: ReadonlySet<string>): OptionExplicitResult
```

- `knownModuleNames` を省略（Pass 1）すると保守的モードで動作し、コール式の bare identifier オブジェクトをすべてスキップする。
- `knownModuleNames` を渡す（Pass 2）と精密モードで動作し、集合に含まれないものは未宣言として検出する。空集合（`new Set()`）を渡した場合はモジュールが存在しない環境と見なし、すべての bare identifier オブジェクトを未宣言として検出する。

---

### checkOptionExplicit が検知できる構文／できない構文

#### 前提：静的解析では識別子の種別が判定できない

VBA の名前解決（§5.6.10）は 6 階層をすべて参照してはじめて「この名前は変数か、定数か、関数か」が確定する。
静的チェッカー（`checkOptionExplicit`）は現在のモジュールの AST しか見ていないため、
他モジュールや型ライブラリの定義が存在するかどうかを知ることができない。

これに加え、VBA には構文上の曖昧さがある。`x = xlUp` の右辺 `xlUp` は：

- 変数への参照（`Dim xlUp As Long` で宣言済み）
- 定数への参照（`Const xlUp = -4162` で宣言済み）
- 引数なし関数の呼び出し（`Function xlUp() As Long` が存在する）

のいずれでも文法的に正しい。VBA では `f` と `f()` が同じ意味を持つため、
コンパイラは名前解決が完了するまで種別を確定できない。

#### パーサーの処理とチェッカーの動作

パーサーは値として現れる裸の識別子を `CallExpression(callee=Identifier('f'), args=[])` として扱う。
チェッカーはこの callee 位置の識別子をスキップする。理由は、`MySub` のような
有効な手続き名を誤って「未宣言変数」と判定しないためである。

一方、構文的に「値でしか使えない位置」に現れる識別子（For ループ変数、関数の引数）は
Identifier ノードとして扱われ、チェッカーの検査対象になる。

| 構文 | Pass 1（単一/マルチ）| Pass 2（マルチのみ）| 理由 |
|------|---------|---------|------|
| `For i = 1 To n`（ループ変数 `i`） | ✅ | ✅ | 構文的に変数しか置けない位置 → Identifier として検査 |
| `SomeFunc(xlUp)`（関数の引数） | ✅ | ✅ | 引数は値位置 → Identifier として検査 |
| `x = xlUp`（代入右辺の裸の識別子） | ❌ | ❌ | 関数呼び出しとも解釈できる → CallExpression の callee は常にスキップ |
| `undeclaredObj.Method()`（コール式のオブジェクト） | ❌ | ✅ | Pass 1 ではモジュール名の可能性があるためスキップ。Pass 2 では `knownModuleNames` と照合し、モジュール名でなければ未宣言と判定 |
| `KnownModule.Proc()`（既知モジュールの修飾呼び出し） | ❌ | ❌ | Pass 2 で `knownModuleNames` に一致 → 正しくスキップ |

---

## 暗黙変数の初期値

`Option Explicit` なしのコードで未宣言変数が参照された場合、`env.get()` は `vbaEmpty`（`null`）で暗黙初期化してその値を返す。VBA 仕様では未宣言変数の初期値は `Empty` であり、数値コンテキストでは `0` として扱われる。

---

## 関連ファイル

- `src/engine/option-explicit-checker.ts` — 静的解析実装
- `src/engine/evaluator.ts` — `callProcedure` の hasVariable チェック、`evaluateModule`、`resolveIdentifiers`、`evaluateConstValue`、`resolveConstIdent`
- `test-libs/test-runner.ts` — `VBARunner` の遅延 `resolveIdentifiers`、`evalVBASingle`/`evalVBAModules`
- `tests/spec/cross-module-const.test.ts` — クロスモジュール定数・循環参照のテスト
