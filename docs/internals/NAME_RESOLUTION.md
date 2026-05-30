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
| 2 モジュール名前空間 | ✅ | `evaluate()` 時にモジュールレベル宣言を global `Environment` へ登録 |
| 3 プロジェクト名前空間 | ⚠️ | モジュール名によるアクセス（`Module.Proc`）は部分対応 |
| 4 同プロジェクト内の他モジュール | ✅ | 複数 `.bas` を `evaluate()` すると global env に蓄積される |
| 5 参照プロジェクト名前空間 | ❌ | 未実装 |
| 6 参照プロジェクトのモジュール内 | ❌ | 型ライブラリなし。`xlUp` 等は未定義 |

tier 6 の定数（`xlUp` 等）はテスト時に `runner.set()` または別モジュールで定義して提供する。

---

## モジュールレベル定数の評価

### 二段階ロード

VBA ではプロジェクト内の全モジュールがコンパイル時に一括解決されるため、他モジュールの `Public Const` を参照する定数はロード順によらず正しく解決される（tier 4）。

当エンジンはこれを二段階ロードで実現する：

```
Pass 1: 全モジュールの evaluate() を実行
  → 手続き・変数・定数が env に登録される

Pass 2: reEvaluateModuleConstsAll() を呼び出す
  → 全モジュールレベル定数を依存グラフでトポロジカルソートして再評価
  → 全モジュールの名前が揃った状態で定数値を確定させる
```

Pass 2 のトポロジカルソートにより、次のどちらのロード順でも `myDir` は `-4162` になる：

```vba
' ExcelConst.bas
Public Const xlUp As Long = -4162

' DataProcessor.bas
Private Const myDir = xlUp   ' xlUp への参照
```

### 循環参照

定数の相互参照はコンパイルエラー（VBA コンパイラの動作）。当エンジンも `topologicalSortConsts` の DFS で閉路を検出した時点でエラーを throw する。

```
Error: Circular reference in constant declarations: constalpha → constbeta → constalpha
```

---

## Option Explicit チェック

### 仕様

§5.6.10 より：名前解決で全階層にマッチしない場合、`explicit-mode`（`Option Explicit` あり）なら「simple name expression is invalid」。

### 実装

チェックは二段階で行う：

1. **静的解析**（`evaluate()` 時）: `checkOptionExplicit` が AST を走査し、未宣言の識別子を使っているプロシージャ名と未宣言名を `optionExplicitViolations: Map<procName, Set<undeclaredNames>>` に記録する
2. **呼び出し時判定**（`callProcedure` 時）: 記録された未宣言名を `env.hasVariable(name)` で再確認する。名前が env に存在すれば tier 1〜4 で解決済みとみなしてエラーを出さない

この設計により、`runner.set()` や別モジュールで提供された名前は、ロード時の静的解析結果によらず呼び出し時に解決済みと扱われる。

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
runner.run('LastRow', [ws]);   // env に xlup なし → Error
```

### checkOptionExplicit が検知できる構文／できない構文

| 構文 | 検知 | 理由 |
|------|------|------|
| `For i = 1 To n`（ループ変数） | ✅ | `checkExprIdents` で Identifier として検査 |
| `SomeFunc(xlUp)`（関数引数） | ✅ | args ループで `checkExpr` が呼ばれる |
| `x = xlUp`（Identifier として解析） | ✅ | `checkExpr` Identifier ケース |
| `x = xlUp`（CallExpression として解析） | ❌ | Bare Identifier callee はスキップ |

Bare Identifier callee をスキップする理由：`f` という名前が変数参照か零引数関数呼び出しかは構文だけでは判別不能なため、手続き名の誤検知を避けている。

---

## env.get() の暗黙初期化

`Environment.get()` は未登録の名前に対して 0 を返すと同時に **env に登録する**：

```typescript
// Environment.get() の末尾
this.variables.set(key, 0);  // 名前が env に登録される
return 0;
```

この挙動により、Pass 1 でモジュールレベルのコードが未定義名を参照した場合、その名前が 0 として env に登録される。Pass 2 の再評価で他モジュールの定数が正しい値に更新されれば問題ない。ただし**どこにも定義されていない名前**は Pass 2 後も 0 のままであり、かつ `env.hasVariable` が true を返すため Option Explicit チェックを通過してしまう。

| 名前の提供方法 | hasVariable | 値 |
|--------------|------------|-----|
| `runner.set('xlup', -4162)` | true | -4162 |
| 別モジュールで `Public Const xlUp = -4162` | true | -4162 |
| どこにも定義されていない（Pass 1 で暗黙登録のみ） | true | 0（不正） |
| どこにも定義されず参照もされていない | false | エラー |

---

## 関連ファイル

- `src/engine/option-explicit-checker.ts` — 静的解析実装
- `src/engine/evaluator.ts` — `callProcedure` の hasVariable チェック、`reEvaluateModuleConstsAll`、`topologicalSortConsts`
- `test-libs/test-runner.ts` — VBARunner の二段階ロード
- `tests/spec/cross-module-const.test.ts` — クロスモジュール定数・循環参照のテスト
- `tests/spec/goto-in-loop.test.ts` — ループ内 GoTo のテスト
