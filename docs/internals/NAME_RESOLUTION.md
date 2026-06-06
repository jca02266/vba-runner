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
| 6 参照プロジェクトのモジュール内 | ⚠️ | 型ライブラリなし。`xlUp` 等は未定義（後述） |

tier 6 の定数（`xlUp` 等）はテスト時に `runner.set()` または別モジュールで `Public Const` として定義して提供する。

### Tier 6 の部分修正（§5.6.10 型名前空間と値名前空間の分離）

VBA では `Dim r As Range` の `Range`（型名前空間）と `Range("A1")` の `Range`（値名前空間）は
別の名前空間に属し、名前が衝突しない。

**修正前の問題**: `setSourceModule('Range')` を呼ぶと `Range` が `VbaNamespaceRef` として
値名前空間（`env.variables`）に登録された。その後 `Range("A1")` を呼ぶと
`OBJECT_REQUIRED (424)` が発生し、`ev.set('Range', fn)` によるモック登録も機能しなかった。

**修正後（`tier6-namespace.test.ts`）**: `VbaNamespaceRef` が値名前空間に残っていても
関数呼び出し経路ではスルーして `SUB_OR_FUNCTION_NOT_DEFINED (35)` へフォールスルーするよう変更。
これにより `ev.set('Range', fn)` が `VbaNamespaceRef` を上書きして `Range("A1")` が動作する。

**残存制限**: 暗黙 Application 経由の解決（`Range("A1")` → `Application.ActiveSheet.Range.Item("A1")`）
は未実装。`Range("A1")` を動作させるにはモック関数を明示的に登録する必要がある（次セクション参照）。

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

---

## 内部実装：データ構造と登録 API

### データ構造の全体像

```
Evaluator
├── env: Environment                        ← 値名前空間
│   ├── variables: Map<string, any>         ← 変数・ビルトイン関数・定数・モジュール名センチネル
│   ├── procedures: Map<string, ProcDecl>   ← VBA Sub / Function / Property 宣言
│   └── types: Map<string, TypeMember[]>    ← Type...End Type で定義した UDT のみ
├── classDefinitions: Map<string, ClassDecl>  ← 型名前空間（Class...End Class）
└── externalObjectFactories: Map<string, ()=>any>  ← CreateObject / New ClassName
```

| データ構造 | 格納されるもの | アクセス経路 |
|---|---|---|
| `env.variables` | 変数・定数・ビルトイン関数・モジュール名センチネル | `env.get(name)` |
| `env.procedures` | VBA Sub / Function / Property 宣言 | `env.getProcedure(name)` |
| `env.types` | UDT メンバー定義 | `env.getType(name)` |
| `classDefinitions` | Class 定義（**型名前空間のみ**、値名前空間とは独立） | `classDefinitions.get(name)` |
| `externalObjectFactories` | `CreateObject` / `New ClassName` のファクトリ | `registerExternalObject()` |

### 値名前空間の解決フロー（`evaluateCallExpression`）

```
1. env.getProcedure(name)          → VBA Sub/Function/Property が見つかれば呼び出す
2. env.get(name) → typeof fn       → JavaScript クロージャ（ビルトイン・ev.set() 注入）なら呼び出す
3. env.get(name) → Array           → 配列インデックスアクセス
4. env.get(name) → __isVbaDict__   → Dictionary(key) アクセス
5. env.get(name) → __isVbaCollection__ → Collection(key) アクセス
6. env.get(name) → __vbaClass__    → デフォルトプロパティ (.Item) アクセス
7. env.get(name) → VbaNamespaceRef → SUB_OR_FUNCTION_NOT_DEFINED へフォールスルー（§5.6.10 修正済み）
8. [Tier 6・未実装] defaultBindingObject[name]  → Application 等のグローバルオブジェクト
9. → SUB_OR_FUNCTION_NOT_DEFINED (error 35)
```

### コンストラクタで自動登録されるもの

#### 標準ライブラリ（`registerStandardLibrary`）

`env.variables` に JavaScript クロージャとして登録。

| 分類 | 登録されるもの |
|---|---|
| プロジェクト参照 | `VBA`（`VbaNamespaceRef`） |
| IO | `MsgBox`（ログ出力スタブ）、`InputBox`（空文字返却スタブ）、`Debug.Print` |
| 型変換 | `CInt` `CLng` `CDbl` `CStr` `CBool` `CDate` … |
| 文字列 | `Len` `Left` `Right` `Mid` `Trim` `UCase` `LCase` `InStr` `Replace` `Format` … |
| 数値 | `Abs` `Int` `Fix` `Rnd` `Round` `Sqr` `Log` `Exp` … |
| 配列 | `Array` `UBound` `LBound` `Join` `Split` … |
| 日付 | `Now` `Date` `Time` `DateAdd` `DateDiff` `Year` `Month` `Day` … |
| 型検査 | `IsEmpty` `IsNull` `IsObject` `IsArray` `IsNumeric` `IsDate` `TypeName` `VarType` |
| エラー | `Err`（`VbaErrObject` インスタンス） |
| vb* 定数 | `vbCrLf` `vbOK` `vbYes` `vbOKCancel` … |
| 環境依存スタブ | `Shell`（実行せずログのみ）、`AppActivate`（no-op） |
| Application | `Application`（`Wait` / `StatusBar` / `DisplayAlerts` / `ScreenUpdating` のみ） |

#### 外部オブジェクト（`registerBuiltinExternalObjects`）

`externalObjectFactories` に登録。`CreateObject("ProgId")` と `New ClassName` の両方から使われる。

| ProgId | クラス名エイリアス | 実装 |
|---|---|---|
| `Scripting.Dictionary` | `Dictionary` | JS `Map` ラッパー |
| `Collection` | — | `VbaCollection` |
| `Scripting.FileSystemObject` | `FileSystemObject` | `MemoryFileSystem` ラッパー |
| `MSXML2.XMLHTTP` / `Microsoft.XMLHTTP` | — | HTTPスタブ |
| `ADODB.Stream` | `Stream` | バイナリ/テキストストリームスタブ |

### ユーザーコードのロード順

```typescript
ev.setSourceModule('Module1');   // 手続きモジュール名を VbaNamespaceRef として env に登録
ev.evaluateModule(ast);          // Pass 1: Sub/Function → env.procedures、変数 → env.variables
                                 //         Class → classDefinitions（env には登録しない）
                                 //         ConstDeclaration はスキップ
ev.resolveIdentifiers([...]);    // Pass 2: 定数の依存グラフ順評価、Option Explicit チェック
```

---

## 外部からの注入 API

### ビルトイン関数のオーバーライド / グローバル関数の追加

```typescript
ev.set('MsgBox', (msg: string) => { captured.push(msg); return 1; }); // vbOK
ev.set('Range', (addr: string) => mockWorksheet.Range(addr));
ev.setConstant('MAX_ROWS', 1000);   // VBA コード側から代入すると Error 5
```

`ev.set()` は `env.variables` に書き込む。`VbaNamespaceRef` が存在しても上書き可能。
**注意**: 同名の VBA Sub/Function が `env.procedures` にある場合は Sub が優先される（`getProcedure` が先）。

### グローバルオブジェクトの注入（ActiveSheet 等）

```typescript
import { MockWorksheet } from '../src/engine/mock/MockWorksheet';

const ws = new MockWorksheet('Sheet1');
ws.setCellValue('A1', 100);
ev.getGlobalEnv().set('ActiveSheet', ws);
// VBA コード: ActiveSheet.Range("A1").Value → 100
```

### CreateObject / New ClassName のモック

```typescript
import { createRegExpMock } from '../../test-libs/regexp-mock';
ev.registerExternalObject('VBScript.RegExp', createRegExpMock);
// factory() の戻り値に __className__ = "RegExp" があれば New RegExp にも自動対応
```

### スパイ（呼び出し記録）

```typescript
const spy = ev.spy('MsgBox');
ev.callProcedure('Main', []);
console.log(spy.calls);        // [[msg1], [msg2], ...]
ev.spy('MsgBox', () => 6);     // vbYes を固定返却してスパイ
```

---

## モックの種類と使い分け

| モックの形式 | 適用場面 | 必要な操作 |
|---|---|---|
| **関数差し替え**（JS クロージャ） | `MsgBox`, `Shell`, `InputBox` など副作用を持つ組み込み | `ev.set(name, fn)` |
| **グローバルオブジェクト注入** | `ActiveSheet`, `ThisWorkbook` など VBA コードが直接参照するオブジェクト | `ev.getGlobalEnv().set(name, obj)` |
| **VBA クラスモック**（`.bas` 拡張構文） | `Worksheet`, `CellRef` など型定義と実装を VBA で書きたい場合 | `setSourceModule` + `evaluateModule` |
| **外部オブジェクトファクトリ** | `CreateObject("VBScript.RegExp")`, `New RegExp` など COM 相当 | `ev.registerExternalObject(progId, factory)` |
| **Tier 6 デフォルト束縛オブジェクト**（未実装） | `Range("A1")` のような修飾なし Excel グローバル関数 | `ev.setDefaultBindingObject(app)`（予定） |

### `.bas` クラスモックの注意点

- **クラス名 ≠ モジュール名** を推奨。`Class Worksheet` を `ExcelMocks.bas` として読み込む構成が衝突を起こさない。
- **クラス名 = モジュール名**（例: `setSourceModule('Range')` + `Class Range`）の場合、
  `Range("A1")` を動かすには `ev.set('Range', fn)` で関数を上書き登録する。

### Tier 6 デフォルト束縛オブジェクト（実装済み）

```
evaluateCallExpression — 名前が見つからない場合の最終フォールバック
│
├─ [早期] 未宣言の名前: env.get() の auto-init より前に defaultBindingObject を確認
│         → Cells(1,1) を 2 回呼んでも 2 回目が OBJECT_REQUIRED にならない
│
├─ [通常の Tier 1〜4 解決]
│
└─ [VbaNamespaceRef の場合] defaultBindingObject も確認（クラスモジュール名の衝突対策）
```

```typescript
import { MockApplication } from '../src/engine/mock/MockExcel';

const app = new MockApplication();
app.ActiveSheet.setCellValue('A1', 42);
ev.setDefaultBindingObject(app);

// VBA: Range("A1").Value  → 42  (CallExpression → Tier 6)
// VBA: Cells(1, 1).Value = 100  (書き込みも可)
// VBA: ActiveSheet.Name  → "Sheet1"  (Identifier → Tier 6)
// VBA: Sheets("Data").Range("B2").Value = 123
```

`MockApplication` は `MockWorksheet` と同列の opt-in モック。登録しない場合は現在と同じ動作（未解決名は `SUB_OR_FUNCTION_NOT_DEFINED`）。

**Option Explicit について**: `Range` / `ActiveSheet` 等は宣言なし識別子として使うため、
これらを使う VBA モジュールは `Option Explicit` なしにする必要がある（実 Excel VBA でも同様）。

---

## 登録優先度と登録タイミング

### 同名が複数箇所にある場合の優先度

| 優先度 | 経路 |
|---|---|
| 1（最高）| VBA Sub/Function/Property（`env.procedures`） |
| 2 | `ev.set()` または `ev.getGlobalEnv().set()` で注入した JS 関数・オブジェクト |
| 3 | `registerStandardLibrary()` が登録したビルトイン（コンストラクタ時） |
| 4 | VBA ローカル変数 / モジュール変数（`env.variables`） |
| — | `classDefinitions`（型名前空間。値名前空間の検索対象外） |
| — | `externalObjectFactories`（`New`/`CreateObject` 時のみ参照） |

### 登録タイミングの制約

| 操作 | `evaluateModule` 前 | 後 | `resolveIdentifiers` 後 |
|---|:---:|:---:|:---:|
| `ev.set(name, fn)` | ✅ | ✅ | ✅ |
| `ev.setConstant(name, val)` | ✅ | ✅ | ✅ |
| `ev.registerExternalObject(...)` | ✅ | ✅ | ✅ |
| `ev.getGlobalEnv().set('ActiveSheet', ws)` | ✅ | ✅ | ✅ |
| `ev.spy(name)` | ⚠️（元の実装がまだない） | ✅ | ✅ |

---

## 関連ファイル

- `src/engine/option-explicit-checker.ts` — 静的解析実装
- `src/engine/evaluator.ts` — `callProcedure` の hasVariable チェック、`evaluateModule`、`resolveIdentifiers`、`evaluateConstValue`、`resolveConstIdent`
- `src/engine/mock/MockWorksheet.ts` — `MockWorksheet` / `MockRange` / `MockApplication`
- `test-libs/test-runner.ts` — `VBARunner` の遅延 `resolveIdentifiers`、`evalVBASingle`/`evalVBAModules`
- `tests/spec/cross-module-const.test.ts` — クロスモジュール定数・循環参照のテスト
- `tests/spec/tier6-namespace.test.ts` — 型名前空間/値名前空間の分離テスト
- `USE_MOCKS.md` — モック実装パターンのユーザー向けガイド
