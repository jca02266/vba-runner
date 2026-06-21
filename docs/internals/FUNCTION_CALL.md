# FUNCTION_CALL.md — 関数コールと引数処理

このドキュメントは `src/engine/evaluator.ts` における「VBA の呼び出し式（`CallExpression`）をどう解決し、
引数をどう束縛して実際の処理に渡すか」を、ユーザー定義 Sub/Function/Property と組み込み関数・モックメソッドの
両方の経路についてまとめる。評価パイプライン全体（Pass 1/Pass 2/Call time）については
[EVALUATION.md](./EVALUATION.md) を参照。

## 2つの世界

VBA の呼び出しには、引数束縛の仕組みが異なる2つの世界がある。

| | 対象 | パラメーター情報の型 | 主な実装場所 |
|---|---|---|---|
| ユーザー定義側 | `Sub`/`Function`/`Property`（標準モジュール・クラス） | `Parameter`（`parser.ts:113`） | `evaluateCallExpression` 内のインライン処理（`evaluator.ts:5879-5954`）、`callClassMethod`（`:3869`） |
| 組み込み・モック側 | 組み込み関数（`Rnd`/`InStr`/`Format`…）、`MockWorksheet.Range` などのモックメソッド | `BuiltinParamSpec` / `BuiltinOverload`（`evaluator.ts:618-639`） | `registerBuiltin`/`registerOverloadedBuiltin`（`:796-845`）+ `resolveCallArgs`（`:3857`） |

両者は完全に独立した型を使うが、引数個数の検証ロジックだけは `ArgBinderParam`（`isOptional`/`isParamArray`
のみを持つ最小形）に変換して `checkArgCountGeneric`（`:3715`）に共有させている。ByRef・`localEnv`・
`defaultValue`（式での既定値）はユーザー定義側だけの概念なので、組み込み側のバインドロジックは意図的に
ユーザー定義側より単純にしてある。

```
                         CallExpression
                               │
                 ┌─────────────┴─────────────┐
                 ▼                            ▼
        env.getProcedure(name)          typeof value === 'function'
        が見つかる（ユーザー定義）       （組み込み関数・モックメソッド・
                 │                       Tier 6 defaultBindingObject）
                 ▼                            ▼
   named/positional 分割 (inline)      resolveCallArgs(fn, argExprs, name)
   checkArgCountGeneric                  ├─ __vbaOverloads__ あり
   ByRef tracking / ParamArray           │    → bindOverloadedCallArguments
   defaultValue 評価                     ├─ __vbaParamSpec__ あり
                 │                       │    → bindCallArguments
                 ▼                       └─ どちらもなし（未移行）
        execProcBody                          → 従来通り位置引数のみ評価
```

## ユーザー定義プロシージャの呼び出し

`evaluateCallExpression`（`evaluator.ts:5814`）の `expr.callee.type === 'Identifier'` 分岐で
`this.env.getProcedure(name)` が見つかった場合の経路。クラスメソッド呼び出し（`obj.Method(...)`、
`obj.__vbaClass__` の分岐）は同じ束縛ロジックを `callClassMethod`（`:3869`）として共有する。

### 1. 名前付き/位置引数への分割

```ts
// evaluator.ts:5879-5888
for (const argExpr of expr.args) {
    if (argExpr.type === 'NamedArgument') {
        const namedArg = argExpr as NamedArgument;
        namedArgs.set(namedArg.name.toLowerCase(), this.evaluateExpression(namedArg.value));
        namedArgExpressions.set(namedArg.name.toLowerCase(), namedArg.value);
    } else {
        positionalArgs.push(this.evaluateExpression(argExpr));
        positionalArgExpressions.push(argExpr);
    }
}
```

`NamedArgument`（`Foo(Bar:=1)`、`parser.ts:466`）は値だけでなく**式そのもの**も
`namedArgExpressions` に保存しておく。これは後段の ByRef 書き戻し（代入先の元の式が必要）のために使う。

### 2. 引数個数の検証

`Parameter`（ユーザー定義）の `isOptional`/`defaultValue` を `ArgBinderParam` の形に変換してから、
組み込み側と共有の `checkArgCountGeneric` に渡す。

```ts
// evaluator.ts:5891-5897
const argBinderParams: ArgBinderParam[] = proc.parameters.map(p => ({
    isOptional: !!p.isOptional || p.defaultValue != null,
    isParamArray: !!p.isParamArray,
}));
this.checkArgCountGeneric(argBinderParams, positionalArgs.length + namedArgs.size);
```

```ts
// evaluator.ts:3715-3728
private checkArgCountGeneric(params: ArgBinderParam[], providedCount: number): void {
    const hasParamArray = params.some(p => p.isParamArray);
    if (hasParamArray) return;

    const maxParams = params.length;
    const minParams = params.filter(p => !p.isOptional).length;

    if (providedCount > maxParams) {
        this.throwVbaError(VbaErrorCode.WRONG_NUMBER_OF_ARGUMENTS, '...'); // 450
    }
    if (providedCount < minParams) {
        this.throwVbaError(VbaErrorCode.ARGUMENT_NOT_OPTIONAL, '...'); // 449
    }
}
```

`ParamArray` を持つ場合は上限チェック自体をスキップする（可変長なので `maxParams` という概念がない）。

### 3. 省略スロット（ギャップ）が必須パラメーターに来ていないかの検証

個数検証（`checkArgCountGeneric`）は「占有スロット数」しか見ないため、`Foo(1, , 3)` のように
`MissingArgument`（カンマだけのギャップ、`parser.ts:2853` で生成）が紛れ込んでいても、合計数が
合っていれば素通りしてしまう。だが省略できるのは `Optional`（または `defaultValue` を持つ）パラメーターだけで、
必須パラメーターの位置にギャップが来るのは VBA 的には不正（コンパイルエラー）な呼び出しである。
これを実行時エラーとして検出するのが `checkNoGapOnRequiredParam`:

```ts
// evaluator.ts:3730-3744
private checkNoGapOnRequiredParam(params: Parameter[], argExprs: Expression[]): void {
    for (let i = 0; i < argExprs.length && i < params.length; i++) {
        if (argExprs[i].type !== 'MissingArgument') continue;
        const p = params[i];
        if (!p.isOptional && p.defaultValue == null && !p.isParamArray) {
            this.throwVbaError(VbaErrorCode.ARGUMENT_NOT_OPTIONAL, 'Argument not optional'); // 449
        }
    }
}
```

`checkArgCountGeneric` の直後（束縛ループの前）に呼ぶことで、束縛ループ自体は「ギャップは常に
Optional/defaultValue 持ちパラメーターの位置にしか来ない」という前提で単純化できる。

このガードは、`MissingArgument` を生む可能性のある呼び出し経路すべて（標準モジュールの
Sub/Function 呼び出し、クラスメソッド呼び出し `obj.Method(...)`、`Implements` 経由のディスパッチ、
デフォルトプロパティ経由の `obj(args)`、モジュール修飾呼び出し `Module1.Proc(...)`）の合計6箇所に
入れている。`callProcedure`（`:2050`）や `callClassMethod`（`:3888`）自身は事前評価済みの
`args: any[]` しか受け取らず、`MissingArgument` ノードかどうかを判定する手段を持たないため、
呼び出し元で `expr.args`（AST レベル）を見ている時点でチェックする必要がある。

### 4. パラメーターへの束縛（ByRef・ParamArray・Optional 既定値）

`proc.parameters` を順に見ながら、各パラメーターの値を「名前付き引数 → 位置引数 → `defaultValue` 式の評価 →
`vbaMissing`」の優先順で決定する（`evaluator.ts:5916-5928`）。

```ts
let argVal: any;
const isMissingSlot = i < positionalArgExpressions.length &&
    positionalArgExpressions[i].type === 'MissingArgument';
if (namedArgs.has(paramNameLower)) {
    argVal = namedArgs.get(paramNameLower);
} else if (i < positionalArgs.length && !isMissingSlot) {
    argVal = positionalArgs[i];
} else if (param.defaultValue) {
    argVal = this.evaluateExpression(param.defaultValue);
} else {
    // checkNoGapOnRequiredParam済みのため、ここに来る時点で param は必ず Optional。
    argVal = vbaMissing;
}
```

以前はこの `else` 分岐が `param.isOptional ? vbaMissing : 0`（`callClassMethod` 側は `vbaEmpty`）と
書かれており、「必須パラメーターの位置にギャップが来た場合のフォールバック値」が実際に式として
存在していた。`checkNoGapOnRequiredParam` の追加によってその分岐は到達不能になったため、
フォールバック式そのものを `vbaMissing` 固定に書き換えて削除している
（`callProcedure`・`callClassMethod` 側の同等コードも同様に簡略化済み）。

`ParamArray` パラメーターは残りの位置引数を丸ごと配列として束縛し、ループを終了する
（`:5903-5910`）。`isByVal` でないパラメーター（既定は ByRef）は `byRefArgs` に
`{ paramName, originalExpr }` として記録し、`execProcBody` 側の戻り時に呼び出し元の変数へ書き戻す
（`ParamArray` 要素も MS-VBAL §5.3.1.5 により ByRef 扱いなので `paramArrayByRefExprs` として別途追跡する）。

束縛が終わると `execProcBody`（`:5956`）にバインド情報一式を渡して本体を実行する。

## 組み込み関数・モックメソッドの呼び出し

### パラメーターメタデータの2つの型

```ts
// evaluator.ts:618-639
export interface BuiltinParamSpec {
    name: string;
    optional?: boolean;
    isParamArray?: boolean;
}

interface ArgBinderParam {
    isOptional: boolean;
    isParamArray: boolean;
}

export interface BuiltinOverload {
    params: BuiltinParamSpec[];
}
```

- **`BuiltinParamSpec[]`** — 「必須引数のあとに Optional 引数が続く」通常形。ユーザー定義 Sub/Function と
  同じ形で、ほとんどの組み込み関数はこれで表現できる。
- **`BuiltinOverload[]`** — VBA 自体にはない、**組み込み関数専用のエンジン内部の機構**。`InStr` のように、
  引数の個数によって先頭の `Start`（Optional）の有無が変わる不規則な関数のためのもの。1つの
  `BuiltinOverload` 内のパラメーターはすべて実質必須で、形をまたいだ Optional は「短い形を別の overload として
  登録する」ことで表現する。

関数は最大でどちらか一方のメタデータしか持たない。メタデータは関数オブジェクトに直接
`__vbaParamSpec__` / `__vbaOverloads__` として生やす（`WeakMap` などは使わない。組み込み関数は freeze
されないため）。

### 登録ヘルパー

```ts
// evaluator.ts:797-815
private registerBuiltin(name: string, fn: any, params: BuiltinParamSpec[], variants: string[] = []) {
    // 登録時セルフチェック: パラメーター名の重複、Optional の後に必須引数が続く形は即 throw
    ...
    (fn as any).__vbaParamSpec__ = params;
    this.envSet(name, fn, variants);
}
```

```ts
// evaluator.ts:824-837
private registerOverloadedBuiltin(name: string, fn: any, overloads: BuiltinOverload[], variants: string[] = []) {
    // 登録時セルフチェック: 1つの overload 内の名前重複、
    // 「アリティとパラメーター名集合が両方一致する overload が2つ以上」は即 throw（曖昧な仕様）
    ...
    (fn as any).__vbaOverloads__ = overloads;
    this.envSet(name, fn, variants);
}
```

どちらも関数の**本体は一切変更しない**。登録の仕方を変えるだけで、引数検証・名前付き引数解決が
後付けで手に入る。括弧無し呼び出しの自動呼び出し可否は、かつて `__vbaAutoCall__` という
専用フラグを関数ごとに手動でセットしていたが（`Rnd`/`DoEvents`/`FreeFile`/`GetObject` などで
何度か付け忘れる実バグがあった）、`registerBuiltin`/`registerOverloadedBuiltin` の導入で
`__vbaParamSpec__`/`__vbaOverloads__` が必ず付くようになったことで、このフラグ自体が冗長になった。
判定に使う情報（必須引数0個か）は `__vbaParamSpec__`/`__vbaOverloads__` に既に入っているため、
別フラグとして事前計算・保持せず、呼び出し時に `isAutoCallable`（`:3890`）でその場で算出する
形に置き換えている。

`InStr`/`InStrB` の実例（`Start` が先頭の Optional）:

```ts
// evaluator.ts:1252-1256
this.registerOverloadedBuiltin('instr', instrFunc, [
    { params: [{ name: 'String1' }, { name: 'String2' }] },
    { params: [{ name: 'Start' }, { name: 'String1' }, { name: 'String2' }] },
    { params: [{ name: 'Start' }, { name: 'String1' }, { name: 'String2' }, { name: 'Compare' }] },
]);
```

`instrFunc` 自身は今まで通り `args.length`/`typeof args[0]` で形を判定するロジックをそのまま持っている。
overload 機構が保証するのは「位置引数だけの呼び出しなら、登録された3つのアリティのどれかでなければ
449/450 を投げる」ことと「名前付き引数を使った場合は該当する overload の順序で位置引数配列を再構築して渡す」
ことだけで、**実際の意味解釈は関数本体が今までどおり担う**。

### 引数の分割・検証・束縛

```ts
// evaluator.ts:3741-3755
private splitCallArgs(argExprs: Expression[]): { namedArgs: Map<string, any>; positionalArgs: any[] } {
    // NamedArgument → namedArgs（小文字キー）、MissingArgument → positionalArgs に undefined、
    // それ以外 → resolveAutoInstance を経て positionalArgs へ
}
```

**`bindCallArguments`**（`BuiltinParamSpec[]` 用、`:3762`）— `checkArgCountGeneric` で個数を検証した後、
各パラメーターを「名前付き引数 → 位置引数 → `undefined`（JS 側のデフォルト引数構文に委ねる）」の順で解決する。
ByRef や `localEnv` を持たない分、ユーザー定義側の束縛ループより単純。

**`bindOverloadedCallArguments`**（`BuiltinOverload[]` 用、`:3795`）の判定は次の2段階:

1. **名前付き引数が1つもない場合**（最も多いケース）— 引数の総数が登録済みのどのアリティにも
   一致しなければ 449/450 を投げる。一致すれば `positionalArgs` を**無加工のまま**返す。
   関数本体の内部判定（`InStr` の `typeof args[0]==='number'` など）が従来通りそのまま動く。
2. **名前付き引数が1つ以上ある場合** — 「アリティが一致」「名前付き引数の名前がすべてその overload の
   パラメーター名に含まれる」「位置引数が占める先頭スロットの名前が名前付き引数と重複しない」
   「残りのスロットがすべて名前付き引数で埋まる」の4条件を満たす overload を絞り込み、ちょうど1つに
   一致したらその overload の順序で位置引数配列を再構築する。0件なら 449/450、2件以上なら
   登録時セルフチェックで本来弾かれているはずの仕様ミスとして内部エラーを投げる（VBA エラーにはしない）。

### `resolveCallArgs` — 呼び出し側の単一エントリーポイント

```ts
// evaluator.ts:3857-3867
private resolveCallArgs(fn: Function, argExprs: Expression[], nameForError: string): any[] {
    const overloads = (fn as any).__vbaOverloads__ as BuiltinOverload[] | undefined;
    if (overloads) return this.bindOverloadedCallArguments(nameForError, overloads, argExprs);
    const spec = (fn as any).__vbaParamSpec__ as BuiltinParamSpec[] | undefined;
    if (spec) return this.bindCallArguments(spec, argExprs);
    return argExprs.map(a => this.resolveAutoInstance(a, this.evaluateExpression(a))); // 未移行関数は従来どおり
}
```

`fn` に何のメタデータも付いていない（まだ `registerBuiltin`/`registerOverloadedBuiltin` に移行していない、
または `env.set` で直接登録された）場合は、従来通り単純な位置引数評価にフォールバックする。これにより
移行は関数単位で安全に進められる。1つの関数の仕様を誤って書いても、影響はその関数だけに留まる。

### 呼び出し箇所（5箇所すべてが `resolveCallArgs` を共有）

`evaluateCallExpression`（`:5814`）内で「VBA の式から JS 関数を呼ぶ」すべての箇所がこの1つのリゾルバーを通る。

| 箇所 | `evaluator.ts` | シナリオ |
|---|---|---|
| Tier 6 早期チェック | `:5983` | `defaultBindingObject`（`MockApplication` 等）のメンバーが関数の場合。`Range("A1")` のように修飾なしで Excel グローバルを呼ぶケース |
| グローバル変数/組み込み関数呼び出し | `:5999` | `this.env.get(name)` が関数を返す場合（組み込み関数、`AddressOf` で取得したコールバック等） |
| `VBA.Func(...)` 修飾呼び出し | `:6089` | `VBA.InStr(...)` のように `VBA` 名前空間越しに組み込み関数を明示的に呼ぶ場合（ユーザー定義の同名関数より常に優先） |
| `obj.Method(...)` メンバー呼び出し | `:6159` | `ws.Range(...)` など、JS オブジェクト（モッククラスのインスタンス）のメソッドを呼ぶ場合 |
| 式の結果を呼ぶ汎用フォールバック | `:6183` | `Array(1, 2)(0)` のように、式の評価結果がそのまま関数である場合 |

`MockWorksheet.Range`/`MockApplication.Range` の2引数対応（`Range(Cell1, Cell2)`）は、まさに
「`obj.Method(...)` メンバー呼び出し」（`:6159`）と「Tier 6 早期チェック」（`:5983`）の2箇所を経由するため、
プロトタイプメソッドに直接 `__vbaOverloads__` を付与するだけで、グローバル組み込み関数とまったく同じ
引数検証・名前付き引数解決を「ただ乗り」で得られる:

```ts
// src/engine/mock/MockExcel.ts:234-237
(MockApplication.prototype.Range as any).__vbaOverloads__ = [
    { params: [{ name: 'Cell1' }] },
    { params: [{ name: 'Cell1' }, { name: 'Cell2' }] },
] satisfies BuiltinOverload[];
```

`registerBuiltin`/`registerOverloadedBuiltin` は `Evaluator` の `private` メソッドなのでモック側からは
呼べないが、メタデータ自体は単なるプロパティなので直接代入すればよい。`resolveCallArgs` は
「`fn` に何が付いているか」だけを見るため、登録元がグローバル組み込み関数かモッククラスかを区別しない。

## 括弧無し呼び出し（auto-call）

`Identifier` の評価（呼び出し式ではなく、単独で参照された識別子）でも「関数だが括弧を付けずに参照した」
ケースを扱う必要がある（`Now`、`Timer`、引数なしの自作 `Function` など）。これは `resolveCallArgs` とは
別の、もう1つの分岐:

```ts
// evaluator.ts:5470-5472
if (typeof v === 'function' && this.isAutoCallable(v)) {
    return v();
}
```

`isAutoCallable`（`:3884`）は `__vbaParamSpec__`/`__vbaOverloads__` からその場で
「必須引数0個か」を判定する。事前計算したフラグを別に持つ必要がないので、登録ヘルパー側
（`registerBuiltin`/`registerOverloadedBuiltin`）はメタデータを付けるだけでよい:

```ts
// evaluator.ts:3884-3893
private isAutoCallable(fn: Function): boolean {
    const overloads = (fn as any).__vbaOverloads__ as BuiltinOverload[] | undefined;
    if (overloads) return overloads.some(o => o.params.length === 0);
    const spec = (fn as any).__vbaParamSpec__ as BuiltinParamSpec[] | undefined;
    if (spec) return !spec.some(p => p.isParamArray) && spec.every(p => p.optional);
    return false;
}
```

ユーザー定義の関数についても同様に、`expr.args` が存在しない（呼び出し式ではない）場合のみ、
必須引数が0個の `Function`/`Property` を自動呼び出しする（`:5473-5485` 付近）。この仕組み自体は
裸の識別子参照（`Identifier` 評価）だけが対象で、`obj.Method` のようなメンバーアクセスには
適用されない。

ただし `obj.Method` 自体は別の経路で括弧無しでも呼ばれる。`evaluateMemberExpression`
（`:6315`）が独立に同等のことをしている。VBA クラスインスタンスなら `Property Get` または
引数0個の `Function` を探して `callClassMethod(obj, ..., [])` で直接呼び（`:6360-6373`）、
JS オブジェクト・モッククラスのプロパティなら `typeof val === 'function' && val.length === 0`
（JS 関数として引数0個）の場合に `val.call(obj)` で呼ぶ（`:6390-6394`）。`col.Count` や
`ws.Name` が括弧無しで動くのはこの経路によるもので、`isAutoCallable` とは無関係の別実装である。

## `spy()` によるメタデータの伝播

テスト用に組み込み関数をラップする `spy(name, returnFn?)`（`:747`）は、元の関数が持つ
`__vbaParamSpec__`/`__vbaOverloads__` をラッパーにもコピーする:

```ts
// evaluator.ts:764-769
if (original && (original as any).__vbaParamSpec__) (spyFn as any).__vbaParamSpec__ = (original as any).__vbaParamSpec__;
if (original && (original as any).__vbaOverloads__) (spyFn as any).__vbaOverloads__ = (original as any).__vbaOverloads__;
```

これを忘れると、スパイした関数だけ引数検証・名前付き引数解決が無効化されてしまう（位置引数の生配列が
そのまま渡る）。組み込み関数を spy するテストを書く際は、この伝播があるおかげで実関数と同じ引数処理を
経由したまま呼び出し記録だけを横取りできる。

## VBA エラーコード

| コード | 定数名 | 発生条件 |
|---|---|---|
| 449 | `ARGUMENT_NOT_OPTIONAL` | 必須引数が足りない（`providedCount < minParams`、または overload の中で最小アリティ未満） |
| 450 | `WRONG_NUMBER_OF_ARGUMENTS` | 引数が多すぎる、または overload のどのアリティにも一致しない（最小アリティ以上だが一致なし） |

両方とも `checkArgCountGeneric`/`bindOverloadedCallArguments` から `this.throwVbaError(...)` で
投げられ、ユーザー定義プロシージャと組み込み関数のどちらでも同じコードが使われる。

## まとめ

| 観点 | ユーザー定義プロシージャ | 組み込み関数・モックメソッド |
|---|---|---|
| パラメーター型 | `Parameter`（`isOptional`/`defaultValue`/`isByVal`/`isParamArray`） | `BuiltinParamSpec` / `BuiltinOverload` |
| 個数検証 | `checkArgCountGeneric`（インライン変換経由） | `checkArgCountGeneric`（`bindCallArguments`/`bindOverloadedCallArguments` 経由） |
| Optional の既定値 | `param.defaultValue` 式を評価 | JS のデフォルト引数構文（呼び出し時は `undefined` を渡すだけ） |
| ByRef | あり（`byRefArgs`/`paramArrayByRefExprs` で書き戻し） | なし |
| 名前付き引数 (`:=`) | パラメーター名と直接マッチ | `BuiltinParamSpec.name` または `BuiltinOverload` 内の名前とマッチ |
| 不規則な形（個数で意味が変わる） | 非対応（VBA 自体にこの概念がない） | `registerOverloadedBuiltin`（エンジン独自機構） |
| 括弧無し自動呼び出し | 必須引数0個の `Function`/`Property` を自動算出 | `isAutoCallable`（`__vbaParamSpec__`/`__vbaOverloads__` からその場で算出） |
| 呼び出しエントリーポイント | `evaluateCallExpression` 内インライン処理 / `callClassMethod` | `resolveCallArgs`（5箇所の呼び出し site で共有） |
