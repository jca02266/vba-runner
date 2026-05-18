# VBA Runner — VBA実行環境 + リファクタリング支援ツール

Excel 不要で VBA コードを実行・テスト・静的解析できる実行環境です。
VBA実行エンジン（パーサー + AST評価器）と、リファクタリング支援のための静的解析CLI（`vba-analyzer`）を備えています。
AIによる支援を前提とした設計で、レガシーVBAコードの品質改善を加速します。

## VBA Runnerの目的
- **Excel 不要での実行・確認**: Excel (Windows/Mac) に依存せず、Node.js 環境で直接 VBA の構文とロジックを実行・検証できます。
- **リファクタリングの支援**: `vba-analyzer` を使って巨大なVBAコードベースの問題箇所（深いネスト・重複ブロック・Excel依存）を特定し、純粋な関数を安全に切り出す検証基盤を提供します。
- **ユニットテストの実行**: TypeScript のテストランナーを通じて、抽出した VBA 関数にモックデータとアサーションを与え、プログラムによるテスト自動化を実現します。

VS Code などのモダンな環境で、AIの支援を受けながら VBA ソースのリファクタリングとテストを進めることができます。

### VSCode拡張機能 (Phase 2)
VBA Runnerにおける第一段階（[MS-VBAL 仕様書](https://learn.microsoft.com/en-us/openspecs/microsoft_general_purpose_programming_languages/ms-vbal/d5418146-0bd2-45eb-9c7a-fd9502722c74)で名前の付いた構文要素・標準ライブラリ関数の実装）は完了しました。仕様書本文に書かれているが個別セクションを持たない **ランタイム挙動** の検証は継続中で、進捗は [`TODO.md` の「VBA ランタイム挙動」](TODO.md#vba-ランタイム挙動)を参照してください。

第二段階として、この実行エンジンを組み込んだ **VSCode拡張機能** を実装しました。`.bas` / `.cls` / `.frm` ファイルを開くと、以下のLSP機能が有効になります。詳細は後述の「[VSCode拡張機能の使い方](#vscode拡張機能の使い方)」を参照してください。

今後の開発方針については [`TODO_NEXT.md`](TODO_NEXT.md) を参照してください。

## 基本方針：環境依存操作のスタブ・仮想化

VBA Runnerは**リファクタリングとユニットテストの実行**を主目的としています。そのため、実行環境（OSやファイルシステム）に副作用を及ぼす、または環境に依存する以下の操作は、原則として**サンドボックス化またはスタブ・フェイク実装**で行います。

- **外部コマンドの実行 (`Shell`)**: 実際のコマンドは実行せず、ログ出力と成功ステータスの返却のみを行います（スタブ化）。
- **ファイル操作 (`Open`, `Kill` 等)**: 指定されたサンドボックスディレクトリ (`sandbox/`) 内に制限されます（仮想化）。
- **環境変数 (`Environ`)**: OSの実環境変数にはアクセスせず、仮想的な環境変数（`.env` 等）を使用します。
- **レジストリ (`GetSetting`, `SaveSetting`)**: Windowsレジストリにはアクセスせず、メモリ上のマップ等による仮想レジストリを使用します。
- **外部オブジェクト・ネットワーク (`CreateObject`, `MSXML2.XMLHTTP`)**: 主要なライブラリをスタブ・フェイク実装として提供し、実際のCOMオブジェクト生成やネットワーク通信は行いません。
- **ユーザー対話 (`MsgBox`, `InputBox`)**: 実行をブロックせず、ログ出力とデフォルト値の返却のみを行います（スタブ化）。

AIを活用する場合、リファクタリングしたいソースコードと同じ階層に VBA Runner を配置し、AIに対して「この `vba-runner` を使ってリファクタリングおよびテストを実行して」と依頼（プロンプト）することで、スムーズに作業を進めることができます。

## VBA言語仕様の拡張機能

### クラス定義の拡張構文

実VBAの標準仕様では、クラスはファイル名（`.cls` ファイル）またはVB_Name属性によってのみ定義されます。`Class...End Class` という構文は存在しません。

本実行エンジンの**拡張機能**として、`.bas` モジュール内に `Class` キーワードを使用してクラスを定義可能にしています。これにより、テストコード内で複数のクラス定義を1つのファイルにまとめることができ、ユニットテスト環境での利便性が向上します。

```vba
' module.bas の例（本エンジンの拡張）
Class MyClass
    Public Name As String
    Sub Initialize(n As String)
        Name = n
    End Sub
End Class

Sub Main()
    Dim obj As MyClass
    Set obj = New MyClass
    obj.Initialize "Test"
    Debug.Print obj.Name
End Sub
```

ただし、**Excelで実行することを想定したVBAコード**（`.cls` ファイル）では、このクラス定義構文を使用しないでください。`.cls` ファイルはVBA標準の形式で、ファイル名がクラス名となります。

## VBA仕様上の制限

本エンジンで実装されているVBAは、[MS-VBAL（MS-VBA Language）仕様 rev.2.4](https://learn.microsoft.com/en-us/openspecs/microsoft_general_purpose_programming_languages/ms-vbal/d5418146-0bd2-45eb-9c7a-fd9502722c74) に準拠しています。以下はVBA言語仕様で定義されている制限事項です。

### モジュール名の長さ制限

VBA仕様により、モジュール名（`.cls` ファイルの場合はファイル名、`.bas` ファイルの場合は名前属性）の最大長は **31文字** に制限されています。（[MS-VBAL 仕様書](https://learn.microsoft.com/en-us/openspecs/microsoft_general_purpose_programming_languages/ms-vbal/d5418146-0bd2-45eb-9c7a-fd9502722c74) §5.2 参照）

```vba
' OK: 31文字以下
Module MyModule                    ' 8文字
Class VeryLongModuleNameExample    ' 27文字

' NG: 32文字以上
' Class ThisIsAVeryLongModuleNameThatExceeds31Chars  ' 計50文字以上
```

## ディレクトリ構成とVBAサンプルコード
VBA Runnerでは、巨大なVBAマクロを「テスト可能な単位」にリファクタリングする事例として、以下の構成を採用しています。

```
project-dir
 ├── vba-runner            ...VBA Runner(`git clone`で配置してください)
 │   ├── src/engine/
 │   ├── sample/src/vba/
 │   ├── sample/src/vba_legacy/
 │   ├── sample/tests/ts/        
 │   ├── tests/spec/             
 │   ├── tests/test-libs-tests/  
 │   └── test-libs/              
 ├── node-vX.X.X-win-x64   ...実行用Node.jsの実行環境（Windows64bitスタンドアローン版の例）
 │   ├── node_modules/        普段Node.jsを使わない人はここにスタンドアローン版を配置するのが手軽です
 │   └── ...
 ├── massive-vba-project   ...リファクタリングしたいVBAソース
 │   ├── Module1.bas
 │   ├── Module2.bas
 │   ├── ...
 │   └── Form1.frm
 └── test/                  ...VBAソースをテストするためのTypeScriptで書かれたテストコード
      ├── Module1Test/
      │   └── test.ts
      └── Module2Test/
          └── test.ts
```

VBA Runnerには、動作確認用のサンプルVBAソースとテストコードが以下のディレクトリに用意されています。これらを使用して、環境のセットアップが正しく完了しているかを確認できます。

- `src/engine/` - TypeScriptで書かれたVBA用Lexer、Parser、Evaluatorのコアエンジン群。
- `sample/src/vba/` - 本番稼働を想定したプロダクションコード。
  - `TaskScheduler.bas`: 処理の見通しが良くなったリファクタリング**後**のメインルーチン。
  - `TaskScheduler_Core.bas`: `TaskScheduler.bas` から抽出された、純粋なビジネスロジック関数モジュール（**ユニットテストの対象**）。
- `sample/src/vba_legacy/` - 参照用のレガシーコード。
  - `TaskScheduler_v1.bas`: リファクタリング**前**の、機能が密結合した巨大なオリジナルソース。
- `sample/tests/ts/` - Node.js 上でVBAスクリプトをAST評価し、自動テストを走らせるTypeScriptスクリプト群。
  - `TaskScheduler_Core.test.ts`: `TaskScheduler_Core.bas` 内の各関数の振る舞いを検証するユニットテスト。
- `test-libs/` - 汎用のテストランナーユーティリティ（VBARunner クラス、assert 等）。

また、以下にVBA評価器自体のテストが配置されています。

- `tests/spec/` - VBAの仕様を確認するための各機能のテスト。
- `tests/test-libs-tests/` - `test-libs/` 配下のテストランナー自身の機能テスト。

## テストの書き方

VBAソースのテストを記述する際は、テストランナー `vba-runner/test-libs/test-runner.ts` を使用します。

> [!TIP]
> 実際にコードを書く前に、ブラウザ上で動作を確認したい場合は [Web UI デモサイト](https://vba-web-runner.netlify.app/) をご利用ください。`Debug.Print` の結果や構文チェックを即座に試すことができます。

### 1. `eval`: VBAの式やコード断片をその場で評価する

VBAの構文をそのまま文字列として渡し、評価結果を取得します。最も手軽にVBAエンジンを試すことができる方法です。

```typescript
import { VBARunner } from '../../test-libs/test-runner';
const vbaRunner = new VBARunner(); // 空の環境を作成

// VBAの計算式を直接評価
const sum = vbaRunner.eval("1 + 2 + 3"); // => 6

// 変数宣言や代入を含む、複数行の処理も実行可能
vbaRunner.eval("Dim x : x = 10 : Debug.Print x * 2");
```

### 2. `run`: ロードしたソース内の関数を実行する

既存のVBAファイル（`.bas`）をロードし、そこに定義されたプロシージャを引数を指定して呼び出します。複雑なビジネスロジックのユニットテストに最適です。

```typescript
// 1. テスト対象のVBAファイルをロード
const vbaRunner = new VBARunner('sample/src/vba/Sample.bas');

// 2. 関数名を指定して実行（第2引数はJavaScriptの配列として引数を渡す）
const result1 = vbaRunner.run('Add', [1, 2]);
const result2 = vbaRunner.run('Multiply', [result1, 2]);

// 3. 結果をアサート
assert.strictEqual(result1, 3);
assert.strictEqual(result2, 6);
```

### 3. 複数ソースの一括ロード

既存のVBAファイル（`.bas`）を格納したディレクトリを指定すれば、配下のソースファイルをすべて読み込みます。
規模の大きなVBAプロジェクトではこの使い方になります。

```typescript
// 1. テスト対象のVBAファイルをロード
const vbaRunner = new VBARunner('sample/src/vba');

// 2. 関数名を指定して実行（第2引数はJavaScriptの配列として引数を渡す）
const result1 = vbaRunner.run('Add', [1, 2]);
const result2 = vbaRunner.run('Multiply', [result1, 2]);

// 3. 結果をアサート
assert.strictEqual(result1, 3);
assert.strictEqual(result2, 6);
```

## テスト環境での `eval` と `run` の使い分け
TypeScript側からVBAのロジックを呼び出す際、用途に応じて以下の2つのメソッドを利用します。

- **`vbaRunner.eval(expressionString)`**:
  - **用途**: VBAの「式（Expression）」や「文（Statement）」を文字列として渡し、その場でVBAとして評価（パース）して実行します。
  - **特徴**: 
    - `AddNumbers(3, 4)` のように、VBAとして有効な式を評価し、その戻り値をダイレクトに取得する場合に向いています。
    - また、`MainLoop` のように括弧なしで「文」として実行されるVBA特有の呼び出しも自動で判別して実行します（この場合の戻り値は `undefined` となります）。
  - **例**: `const result = vbaRunner.eval("AddNumbers(3, 4)");` または `vbaRunner.eval("MainLoop");`

- **`vbaRunner.run(procedureName, args)`**:
  - **用途**: VBAのサブルーチン（`Sub`）や関数（`Function`）に、TypeScript側からの配列引数を渡して実行します。
  - **特徴**: トップレベルから直接プロシージャ名を指定して呼び出すための「標準的」な実行メソッドです。
  - **例**: `const result = vbaRunner.run("AddNumbers", [100, 200]);`

#### Sub と Function の戻り値の対応表

`run` と `eval` はいずれもプロシージャを実行し戻り値を取得できますが、引数の渡し方に違いがあります。

| 呼び出し対象 | `run` の書き方（例） | `eval` の書き方（例） | 戻り値 |
| :--- | :--- | :--- | :--- |
| **Sub** | `run("Log", [10, "msg"])` | `eval("Log 10, ""msg""")` | `undefined` |
| **Function** | `run("Add", [1, 2])` | `eval("Add(1, 2)")` | 戻り値 |
| **VBA式** | (非対応) | `eval("1 + 2 + Add(1, 2)")` | 評価・計算結果 |
| **VBA文** | (非対応) | `eval("Dim x : x = 10 : Debug.Print x")` | `undefined` |
| **変数評価** | (非対応) | `eval("ActiveSheetName")` | 変数の現在の値 |

- **`run` の強み**: TypeScript 側の変数（配列やオブジェクト、数値など）を、そのまま第 2 引数の配列として安全に渡せます。
- **`eval` の強み**: `vbaRunner.eval("Add(1, 2) * 10")` のように、戻り値を利用した計算式や複数行の処理を、VBA の記法でそのまま記述・評価できます。

> **Tip**: 基本的には、引数があるプロシージャのテストには `run` を使用します。サブルーチン（`Sub`）を実行した後に、その影響で書き換わった**グローバル変数やモジュール変数の値を個別にアサート（確認）したい場合**などは、`eval("VariableName")` を組み合わせて使うのが定石です。
> また、長期的には `run` だけで完結できるような「副作用のない純粋な関数（Function）」にロジックを切り出し、リファクタリングしていくことが推奨されます。

## テストコードで使う固有型と値

VBAの型システムはJavaScriptと異なるため、`run` / `eval` の戻り値をアサートする際に以下の対応表を参照してください。

### スカラー値の対応

| VBA の値          | JavaScriptの型・値    | 説明                                                       |
|-------------------|-----------------------|------------------------------------------------------------|
| `Empty`           | `null`（`vbaEmpty`）  | 未初期化のVariant型変数。型指定なしの変数は`vbaEmpty`で初期化。`evaluator.ts`の`vbaEmpty`定数と同値 |
| `Nothing`         | `null`                | 未設定のオブジェクト参照                                   |
| `True`            | `VbaBoolean`オブジェクト | 詳細は[制限事項](#制限事項)を参照                         |
| `False`           | `VbaBoolean`オブジェクト | 詳細は[制限事項](#制限事項)を参照                         |
| 数値              | `number`              | 詳細は[制限事項](#制限事項)を参照                         |
| 文字列            | `string`              | ダブルクォート除去済みの生文字列                           |
| `Null`            | `Symbol`（`vbaNull`） | `vbaNull`定数と同値                                       |

```typescript
import { vbaEmpty } from '../../src/engine/evaluator';

assert.strictEqual(result, vbaEmpty);   // VBAのEmpty（null）と比較
assert.strictEqual(result, null);       // 同上（直接nullでも可）
```

### ブール値のアサーション

VBAの比較演算や `IsNull`、`CBool` などの関数は `-1`（True）または `0`（False）を返します。
`assert.isTrue` / `assert.isFalse` を使うと、`vbaTrue`/`vbaFalse` の定数をインポートせずに簡潔にアサートできます。

```typescript
import { assert } from '../../test-libs/test-runner';

assert.isTrue(vbaRunner.eval('1 = 1'),        '等値比較');
assert.isFalse(vbaRunner.eval('IsNull(123)'), 'IsNull(123)');

// assert.strictEqual で書く場合は vbaTrue/vbaFalse をインポートする必要がある
// import { vbaTrue, vbaFalse } from '../../src/engine/evaluator';
// assert.strictEqual(result, vbaTrue);
```

#### `assert.isTrue` / `assert.isFalse` の比較ルール

`assert.isTrue` / `assert.isFalse` は **VBA の `= True` / `= False`（厳密比較）と等価** です。
本エンジン内部では VBA Boolean を `vbaTrue` / `vbaFalse` のシングルトンで表すため、これらの assert はシングルトンと一致するかを **リファレンス比較** します。

```typescript
assert.isTrue(vbaTrue);          // ✅ OK
assert.isTrue(vbaRunner.eval('1 = 1'));  // ✅ OK（比較演算は vbaTrue を返す）

assert.isTrue(-1);               // ❌ FAIL（数値 -1 は vbaTrue シングルトンとは別物）
assert.isTrue(vbaRunner.eval('Not 5')); // ❌ FAIL（Not <数値> はビット反転の数値 -6 を返す）
```

「数値が **真とみなせる**（truthy）こと」を確認したいときは `assert.ok` を使います。

```typescript
assert.ok(vbaRunner.eval('5'));       // ✅ OK（5 は truthy）
assert.ok(vbaRunner.eval('Not 5'));   // ✅ OK（-6 は truthy）
assert.ok(vbaRunner.eval('0'));       // ❌ FAIL（0 は falsy）
```

#### VBA の Truthy/Boolean 仕様（補足）

VBA の Boolean は **値そのものは -1 (True) / 0 (False) のみ** ですが、文脈によって解釈が変わります。

| 文脈 | 例 | 結果 |
|------|-----|------|
| **代入時の暗黙型変換** | `Dim b As Boolean : b = 5` | `b = -1`（非 0 → True に自動変換） |
| **`If` 条件の Boolean 評価** | `If 5 Then ...` | 実行される（非 0 は True 扱い） |
| **比較演算 `= True` / `= False`** | `5 = True` | **False**（数値 5 と -1 は等しくない） |

### `Not` 演算子の戻り値

`Not` は引数の型によって戻り値の型が変わります（[MS-VBAL](https://learn.microsoft.com/en-us/openspecs/microsoft_general_purpose_programming_languages/ms-vbal/d5418146-0bd2-45eb-9c7a-fd9502722c74) §5.6.9.8.1）:

- **Boolean オペランド** → Boolean 反転（シングルトン）
  - `Not vbaTrue` → `vbaFalse`
  - `Not vbaFalse` → `vbaTrue`
- **数値オペランド** → ビット反転（数値）
  - `Not 5` → `-6`
  - `Not 0` → `-1`（ただし数値の `-1` であって Boolean シングルトンではない）

そのため、`Not <数値>` の結果を **Boolean として** assert したい場合は、VBA 側で `CBool(...)` を挟むか、TypeScript 側で `assert.ok(...)` を使ってください。

### 配列

本実行エンジンの配列表現の基本ルール：**VBA インデックスと JavaScript の配列インデックスを一致させる**。
JavaScript 配列の `length` は常に `upper + 1`、`[0]`〜`[lower-1]` は未使用（`undefined`）、`[lower]`〜`[upper]` に値が入ります。

#### デフォルト (`Option Base 0`) — `Dim arr(N)`

`Dim arr(N)` は `arr(0)`〜`arr(N)` の `N + 1` 要素配列。JS 側でも `arr[0]`〜`arr[N]` がそのまま有効です。

**VBA側:**
```vba
Function GetArray()
    Dim arr(5)
    arr(1) = "first"
    GetArray = arr
End Function
```

**TypeScript側:**
```typescript
const arr = vbaRunner.run('GetArray', []) as any[];
assert.strictEqual(arr.length, 6);    // N=5 のため要素数は 6
assert.strictEqual(arr[1], "first");
```

#### `Option Base 1` — `Dim arr(N)`

モジュール先頭で `Option Base 1` を宣言すると、`Dim arr(N)` は `arr(1)`〜`arr(N)` になります。JS 配列も `length = N + 1` のまま `[0]` は未使用となり、`[1]`〜`[N]` に値が入ります。

**VBA側:**
```vba
Option Base 1

Function GetArray()
    Dim arr(3)
    arr(1) = "a"
    arr(3) = "c"
    GetArray = arr
End Function
```

**TypeScript側:**
```typescript
const arr = vbaRunner.run('GetArray', []) as any[];
assert.strictEqual(arr.length, 4);          // [0]〜[3]
assert.strictEqual(arr[0], undefined);      // Option Base 1 のため [0] は未使用
assert.strictEqual(arr[1], "a");            // VBA: arr(1)
assert.strictEqual(arr[3], "c");            // VBA: arr(3)
```

#### `Dim arr(L To U)` / `ReDim arr(L To U)`（任意の下限）

明示的に下限を指定した場合も同じルールです。`length = U + 1`、`[0]`〜`[L-1]` は未使用、`[L]`〜`[U]` が有効。

**VBA側:**
```vba
Function GetArr1To3() As Variant
    Dim arr(1 To 3) As Double
    arr(1) = 0.5
    arr(2) = 1.0
    arr(3) = 1.5
    GetArr1To3 = arr
End Function
```

**TypeScript側:**
```typescript
const arr = vbaRunner.run('GetArr1To3', []) as any[];
assert.strictEqual(arr.length, 4);          // [0]〜[3]、[0] は未使用
assert.strictEqual(arr[0], undefined);      // VBA インデックス外
assert.strictEqual(arr[1], 0.5);            // VBA: arr(1)
assert.strictEqual(arr[2], 1.0);            // VBA: arr(2)
assert.strictEqual(arr[3], 1.5);            // VBA: arr(3)
```

### Scripting.Dictionary

`CreateObject("Scripting.Dictionary")` が返すオブジェクトの構造:

```typescript
interface VbaDict {
    __isVbaDict__: true;
    __map__: Map<string, any>;  // 内部のJS Map（テストでの直接参照用）
    add(key: string, value: any): void;
    exists(key: string): number;  // vbaTrue(-1) または vbaFalse(0)
    items(): any[];
    keys(): string[];
}
```

テストコードでは `__map__` を通じて内容を検証できます:

```typescript
const dict = vbaRunner.run('BuildDict', []) as any;
assert.strictEqual(dict.__map__.get('key1'), 'value1');
assert.isTrue(dict.exists('key1'));
```

### CreateObject のファクトリ差し替え（モック登録）

VBA の `CreateObject(progId)` が返すオブジェクトを **テストコード側でモックに差し替え** できます。これにより、本体（`src/engine/evaluator.ts`）に手を入れずに、テスト固有の COM オブジェクトをエミュレートできます。

#### API
- `Evaluator.registerExternalObject(progId, factory)`
- `VBARunner.registerExternalObject(progId, factory)` （上記のプロキシ）

`factory` は `CreateObject` が呼ばれるたびに **新しいオブジェクト** を返す関数です。登録されたファクトリは既存の組み込みスタブ（`Scripting.Dictionary` 等）よりも優先されます。

#### モック実装の作り方

VBA 側の API に合わせて以下を実装します。

- **プロパティ**: 通常の JS オブジェクトのフィールドとして書きます。Boolean は `vbaTrue / vbaFalse`（`-1 / 0`）を使う。
- **メソッド**: 関数フィールドとして書きます（VBA からは `obj.Method(args)` で呼ばれる）。
- **デフォルトプロパティ（`obj(i)` 形式）**: オブジェクト本体を **関数** にして、追加のプロパティを生やします（例: `SubMatches(i)` や Collection の `Item(i)`）。
- **`For Each` 対応**: 関数 / オブジェクトに `[Symbol.iterator]` を実装する。
- **参照設定 (`Dim x As ClassName` / `New ClassName`) 対応** *(任意)*: factory が返すオブジェクトに `__className__: 'クラス名'` を含めると、その名前でも自動的に別名登録されます。

#### サンプル: `VBScript.RegExp` のモック

`test-libs/regexp-mock.ts` に実装例があります（JS の `RegExp` をベースに VBA RegExp 互換 API を提供）。

```typescript
// test-libs/regexp-mock.ts より抜粋
import { vbaTrue, vbaFalse } from '../src/engine/evaluator';

export function createRegExpMock(): any {
    const state: any = {
        __isVbaRegExp__: true,
        __className__: 'RegExp',   // ★ 参照設定相当: New RegExp / Dim x As RegExp で使う名前
        pattern: '',
        ignorecase: vbaFalse,
        global: vbaFalse,
        multiline: vbaFalse,

        test(s: any) { /* JS RegExp.test を呼んで vbaTrue/vbaFalse を返す */ },
        execute(s: any) { /* MatchCollection を返す */ },
        replace(s: any, repl: any) { /* JS String.replace を呼ぶ */ },
    };
    return state;
}
```

#### テストコードでの登録例

```typescript
import { VBARunner, assert } from '../../test-libs/test-runner';
import { createRegExpMock } from '../../test-libs/regexp-mock';

const vbaRunner = new VBARunner('macro.bas');
// 1 回の登録で「CreateObject」も「New / Dim As」も同じ factory が使われる
vbaRunner.registerExternalObject('VBScript.RegExp', createRegExpMock);

const result = vbaRunner.run('CountDigits', ['abc 123 def 456']);
assert.strictEqual(result, 2);
```

VBA 側のコード例 — どちらの書き方も同じ結果になります：

**`CreateObject` 経由（progId 文字列）:**
```vba
Function CountDigits(s As String) As Long
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "\d+"
    re.Global = True
    CountDigits = re.Execute(s).Count
End Function
```

**参照設定相当（`New` 演算子）:**
```vba
Function CountDigits(s As String) As Long
    Dim re As RegExp           ' 型として直接参照
    Set re = New RegExp        ' New 演算子でインスタンス化
    re.Pattern = "\d+"
    re.Global = True
    CountDigits = re.Execute(s).Count
End Function
```

#### 「参照設定」相当が動く仕組み

- `registerExternalObject(progId, factory)` は登録時に factory を 1 度呼び、戻り値の `__className__` プロパティを読み取ります。
- 値があれば、その class 名（小文字化）でもファクトリを登録します。
- VBA の `New ClassName` / `Dim x As ClassName` は内部的に同じ class 名 lookup を行うため、`CreateObject` と同じ factory が使われます。
- `__className__` を持たないオブジェクトを返す factory は `CreateObject` 経由でのみアクセスできます。

#### 組み込み外部オブジェクトも参照設定相当をサポート

本実行エンジンに同梱されている組み込み外部オブジェクトは、すべて `registerExternalObject` 経由で登録されているため、`CreateObject` と `New` の両方の構文に対応しています。

| ProgID | 参照設定相当の class 名 |
|--------|---------------------|
| `Scripting.Dictionary` | `Dictionary` |
| `Scripting.FileSystemObject` | `FileSystemObject` |
| `MSXML2.XMLHTTP` / `Microsoft.XMLHTTP` | `XMLHTTP` |
| `ADODB.Stream` | `Stream` |

例:
```vba
' どちらも同じ Dictionary オブジェクトを得る
Dim d1 As Object
Set d1 = CreateObject("Scripting.Dictionary")

Dim d2 As New Dictionary

Dim d3 As Dictionary
Set d3 = New Dictionary
```

テストでこれらをモックに差し替えたい場合は、同じ progId / class 名で `registerExternalObject` を呼ぶと **組み込みの実装を上書き** できます。

#### Auto-Instantiation 仕様（`Dim x As New ClassName`）

VBA の `Dim x As New ClassName` は通常の宣言とは異なる特殊な挙動を持ち、本実行エンジンもこれに準拠しています：

- **遅延インスタンス化**: 宣言時点では実際のインスタンスは作成されず、最初のメンバ参照やメソッド呼び出しで自動生成される
- **`Nothing` 後の自動再生成**: `Set x = Nothing` した後でも、再度参照すると新しいインスタンスが自動生成される
- **`x Is Nothing` は常に False**: auto-instance 変数は参照前でも `Nothing` ではないと判定される

```vba
Dim a As New Dictionary  ' この時点ではインスタンスは未生成
Debug.Print a Is Nothing ' → False（auto-instance なので Nothing ではない）

a.Add "x", 1             ' 初回アクセスで生成される
Debug.Print a.Count      ' → 1

Set a = Nothing          ' リセット
Debug.Print a Is Nothing ' → False（再アクセス対象なので Nothing 扱いではない）
a.Add "y", 2             ' 自動再インスタンス化 → 新しい Dictionary
Debug.Print a.Count      ' → 1（"x" は前のインスタンスにあったので新インスタンスにはない）
```

通常の `Dim x As ClassName`（`As New` なし）はオブジェクト型参照の標準的な挙動で、デフォルトは `Nothing` です：

```vba
Dim b As Dictionary       ' b は Nothing
Debug.Print b Is Nothing  ' → True
b.Add "x", 1              ' エラー（オブジェクト参照が未設定）

Set b = New Dictionary    ' 明示的にインスタンス化
b.Add "x", 1              ' OK
Set b = Nothing
b.Add "y", 2              ' エラー（再アクセスでも自動生成されない）
```

> **動作確認**:
> - `tests/test-libs-tests/regexp-mock.test.ts` — ユーザー登録モック（VBScript.RegExp）の `CreateObject` 形式と `New RegExp` 形式
> - `tests/spec/createobject.test.ts` — 組み込みオブジェクト（Dictionary, FSO, XMLHTTP, ADODB.Stream）の `New` 構文サポート、Auto-Instantiation 仕様
>
> 同じ仕組みで `MSXML2.XMLHTTP` や自社の独自 COM オブジェクトもモック化できます。

---

## Excel オブジェクトのモック化（Sheets / Range / Application）

VBA マクロの中でも **Excel オブジェクトへの依存**（`Sheets(...)`, `Range(...)`, `Application.WorksheetFunction` など）は頻出です。これらはこのエンジン上に Excel は存在しないため、テスト時にはモックが必要です。

### 推奨アプローチ：依存の分離を優先する

モック化の前に、まず **Excel I/O とドメインロジックを分離**することを強く推奨します（詳細は [`docs/TESTING_STRATEGY.md`](docs/TESTING_STRATEGY.md) / [`docs/REFACTORING_GUIDE.md`](docs/REFACTORING_GUIDE.md)）。

```vba
' ❌ テストしにくい：ロジックが Excel I/O に直結している
Sub CalculateSalesTotal()
    Dim total As Long
    total = Sheets("Data").Range("A1").Value + Sheets("Data").Range("B1").Value
    Sheets("Result").Range("A1").Value = total
End Sub

' ✅ テストしやすい：ロジックを純粋関数として分離
Function CalcTotal(a As Long, b As Long) As Long
    CalcTotal = a + b
End Function
```

### MockApplication / MockWorksheet を使う

VBA Runner には `src/engine/mock/MockWorksheet.ts` として `MockApplication` と `MockWorksheet` が実装済みです。`Sheets(...)` のような Excel 依存を分離しきれない場合にモックを使います。

```typescript
import { MockApplication } from '../../src/engine/mock/MockWorksheet';
import { Lexer } from '../../src/engine/lexer';
import { Parser } from '../../src/engine/parser';
import { Evaluator } from '../../src/engine/evaluator';

const mockApp = new MockApplication();
mockApp.Sheets('Data').setCellValue('A1:A3', [[10], [20], [30]]);

const ev = new Evaluator(() => {});
ev.evaluate(new Parser(new Lexer(vbaCode).tokenize()).parse());

// Sheets 関数をモックに置き換えてエンジンに注入
ev.getGlobalEnv().set('Sheets', (name: string) => {
    const ws = mockApp.Sheets(name);
    return { Range: (addr: string) => ws.Range(addr) };
});

const result = ev.callProcedure('GetTotal', []);
```

### MockApplication の拡張

前述のとおり、ドメインロジックの分離を優先しているため、`MockApplication` などのモックオブジェクトは機能が限定的です。それでも既存モックを拡張することで当面の問題に対処できます。

ここでは `WorksheetFunction` のモック化を例に説明します。`WorksheetFunction` は Excel のワークシート関数（`VLookup` など）を VBA から呼び出す機能であり、汎用的なモックをあらかじめ用意するのが難しい典型例です。必要な関数だけを継承先に実装する方法が実用的です。

**`MockApplication` を継承**して `worksheetfunction` プロパティを追加します。`application` オブジェクトを丸ごと差し替えると `Sheets` などの既存機能が失われるため、継承による拡張を使います。

> **注意**: エンジンは VBA の識別子を小文字で解決します（`Application.WorksheetFunction` → `application.worksheetfunction`）。プロパティ名はすべて lowercase で定義してください。

```typescript
import { MockApplication } from '../../src/engine/mock/MockWorksheet';

// MockApplication を継承して WorksheetFunction を追加
class MockApplicationEx extends MockApplication {
    readonly worksheetfunction = {
        vlookup: (lookupVal: any, table: any[][], col: number) => {
            const row = (table as any[][]).find(r => r[0] === lookupVal);
            if (!row) throw Object.assign(new Error('N/A'), { number: 1004 });
            return row[col - 1];
        },
        sum: (...args: any[]) =>
            args.flat(Infinity).reduce((a: number, b: any) => a + Number(b), 0),
        max: (...args: any[]) => Math.max(...(args.flat(Infinity).map(Number))),
        countif: (range: any[][], criteria: any) =>
            (range as any[][]).flat().filter(v => v === criteria).length,
    };
}

// テストデータを準備
const mockApp = new MockApplicationEx();
mockApp.Sheets('Data').setCellValue('A1:B3', [
    ['Alice', 90000],
    ['Bob',   75000],
    ['Carol', 85000],
]);

// エンジンに注入
ev.getGlobalEnv().set('Sheets', (name: string) => {        // Sheets("Data") グローバル呼び出し用
    const ws = mockApp.Sheets(name);
    return { Range: (addr: string) => ws.Range(addr) };
});
ev.getGlobalEnv().set('application', mockApp);             // Application.WorksheetFunction 用
ev.getGlobalEnv().set('worksheetfunction', mockApp.worksheetfunction); // WorksheetFunction 省略形用
```

#### VBA 側の呼び出しパターン

Excel VBA では `Application` のパブリックメンバーはグローバルスコープに自動公開されるため、以下の **2 つの書き方が等価**です。上記の注入でどちらも動作します。

```vba
Function GetSalary(name As String) As Long
    Dim data As Variant
    data = Sheets("Data").Range("A1:B3").Value

    ' パターン 1: Application. あり（ev に application を注入）
    GetSalary = Application.WorksheetFunction.VLookup(name, data, 2)

    ' パターン 2: Application. 省略（ev に worksheetfunction を直接注入）
    ' GetSalary = WorksheetFunction.VLookup(name, data, 2)
End Function
```

> **詳細は [`docs/MOCK_GUIDE.md`](docs/MOCK_GUIDE.md) を参照してください。**
> MockWorksheet の API リファレンス・使用例・自作モックの書き方・Part 1〜3 の段階的ガイドが含まれています。

---

## VBA 実行時の終了ステータスと例外の検知
 
TypeScript 側から `vbaRunner.run()` または `vbaRunner.eval()` を呼び出した際、VBA の終了状態（正常終了、`End` による強制終了、実行時エラー）を以下のように区別して検知できます。
 
### 1. 正常終了（Sub/Function の戻り値）
 
プロシージャが最後まで実行された、または `Exit Sub/Function` で抜けた場合、戻り値が TypeScript 側に返ります。
 
```typescript
try {
    const result = vbaRunner.run("MyFunction", [10]);
    console.log("正常終了:", result);
} catch (e) {
    // 正常終了時はここには来ない
}
```
 
### 2. `End` ステートメントによる強制終了
 
VBA コード内で `End` ステートメントが実行されると、呼び出し側に `{ type: 'Terminate' }` というオブジェクトがスローされます。これは通常のランタイムエラーとは区別してハンドルできます。
 
```typescript
try {
    vbaRunner.eval("End");
} catch (e: any) {
    if (e && e.type === 'Terminate') {
        console.log("End ステートメントによりプログラムが終了しました");
    }
}
```
 
> **設計意図**: VBA の `End` は、すべてのコールスタックを破棄し実行を即座に中断する強力な命令です。本エンジンではこの「強制終了」の振る舞いを確実に再現するため、JavaScript の例外機構を利用してトップレベルまで制御を戻す設計としています。通常の `Exit Sub/Function` による正常な脱出とは明確に区別されます。
 
### 3. 実行時エラー（例外）
 
構文エラー（パース時）や、未定義の関数呼び出し、ゼロ除算、型不一致などの実行時エラーが発生した場合は、標準的な `Error` オブジェクト（またはメッセージ文字列を含むオブジェクト）がスローされます。
 
```typescript
try {
    vbaRunner.run("NonExistentSub", []);
} catch (e: any) {
    console.error("実行時エラー:", e.message); // 例: "Procedure not found: NonExistentSub"
}
```
 
### 4. `Stop` ステートメント
 
`Stop` ステートメントは、現在の実装では `console.log` に `STOP Statement encountered` と出力するのみで、プログラムの実行自体はそのまま継続されます（デバッガのブレークポイントのような動作は行いません）。
 

### 未初期化変数の暗黙初期化

VBA仕様に従い、宣言されただけで値を設定していない変数は `0` として扱われます（`vbaEmpty`/`null` ではありません）。テストで「初期化されていないこと」を確認したい場合は実際の返り値（`0` または `""`）を確認してください。

```typescript
// VBA: Dim x As Long （代入なし）
assert.strictEqual(vbaRunner.run('GetX', []), 0);
```

## ファイル入出力のSandbox方針

VBAコードからのファイル入出力はすべて**Sandboxディレクトリ**内に限定します。これにより、実行環境のファイルシステムを保護しつつ、VBAのファイル操作命令をテスト可能にします。

### Sandboxの動作ルール

1. **Sandboxルートの指定**: Evaluatorの初期化時にSandboxルートディレクトリを指定します。デフォルトは `sandbox/`（リポジトリルートからの相対パス）です。

2. **Windowsパスの仮想化**: VBAコード内の絶対パスはドライブレターをサブディレクトリとしてSandboxルート配下にマッピングします。
   - `C:\foo\bar.txt` → `{sandboxRoot}/c/foo/bar.txt`
   - `D:\data\output.csv` → `{sandboxRoot}/d/data/output.csv`
   - ドライブレター（`C:`、`D:` など）は小文字に正規化してSandboxルート配下のディレクトリとして扱います。

3. **上位ディレクトリへのアクセス禁止**: `../` などを使ってSandboxルートの外へのアクセスを試みた場合はランタイムエラーになります。

4. **相対パスの扱い**: VBAコード内の相対パスはSandboxルートを起点として解決されます。

### パス変換ユーティリティ（テスト用）

`test-libs/sandbox.ts` に実パスと仮想パス（VBAから見えるWindowsパス）を相互変換するユーティリティを提供します。

```typescript
import { SandboxPath } from '../../test-libs/sandbox';

const sandbox = new SandboxPath('./sandbox');

// 実パス → VBA上の仮想パス（Windows形式）
sandbox.toVirtualPath('./sandbox/c/foo/bar.txt'); // => "C:\\foo\\bar.txt"

// VBA上の仮想パス → 実パス
sandbox.toRealPath('C:\\foo\\bar.txt');             // => "./sandbox/c/foo/bar.txt"
sandbox.toRealPath('C:/foo/bar.txt');               // => "./sandbox/c/foo/bar.txt"
sandbox.toRealPath('D:/data/output.csv');           // => "./sandbox/d/data/output.csv"
```

テストコードでは、VBAコードに渡すパスをこのユーティリティで組み立てることで、実ファイルシステムへの影響を限定しながらファイル操作の動作を検証できます。

### 環境変数のSandbox化

VBAの `Environ` 関数による環境変数アクセスも実際のOSの環境変数とは切り離します。

- **定義方法**: Sandboxの設定ファイル（`{sandboxRoot}/.env`）に `KEY=VALUE` 形式で記述するか、Evaluatorの初期化時にメモリ上のマップとして渡します。
- **実環境変数へのアクセス禁止**: VBAコードから `Environ("PATH")` などを呼び出しても、OSの実際の環境変数は参照されません。Sandbox内に定義のないキーは空文字列 `""` を返します。
- **優先順位**: メモリ上の定義 → `{sandboxRoot}/.env` の順に参照します。

```typescript
// メモリ上で環境変数を定義する例
const sandbox = new SandboxPath('./sandbox', {
  env: {
    USERNAME: 'testuser',
    APPDATA: 'C:\\Users\\testuser\\AppData\\Roaming',
  }
});
```

## 仮想ファイルシステム (VFS) の利用と将来構想

VBA Runnerでは、ブラウザ上での動作や、テストの安全性を高めるために **仮想ファイルシステム (VFS)** を導入しています。

### 将来構想：VFS を基本とした実行環境
長期的には、Node.js 環境においても実際のディスクを操作する `NodeFileSystem` ではなく、メモリ上で完結する `MemoryFileSystem` (VFS) をデフォルトの実行環境にすることを構想しています。これにより、以下のメリットが得られます：
- **安全性**: テスト実行によってホスト OS のファイルが誤って削除・上書きされるリスクをゼロにします。
- **再現性**: テスト実行ごとにクリーンなファイルシステム状態から開始でき、並列実行も容易になります。
- **ブラウザ互換性**: テストコードそのものをブラウザ上でそのまま実行可能になります。

### VFS の有効化方法
`VBARunner` の初期化時に `useVirtualFS: true` を指定するか、環境変数 `USE_VFS=1` を設定することで VFS モードになります。

```typescript
// 個別にVFSを有効化
const vbaRunner = new VBARunner('source.bas', { useVirtualFS: true });

// または環境変数で一括切り替え
// USE_VFS=1 node tests/spec/xxx.test.cjs
```

### JavaScript から VFS への直接アクセス (テストデータの準備)

テストの準備や実行結果の検証のために、JavaScript から VFS に対して直接ファイルを読み書きできます。

**簡潔な例**（VFS と Node.js 両対応）:

```typescript
import { VBARunner } from './test-libs/test-runner';

// VFS モード
const vbaRunner = new VBARunner('source.bas', { useVirtualFS: true });
const fs = vbaRunner.evaluator.fs;
fs.writeFileSync('/sandbox/c/input.txt', "データ");
vbaRunner.run('ProcessFile', []);
const result = fs.readFileSync('/sandbox/c/output.txt', 'utf-8');
```

```typescript
// Node.js モード (実ファイルを使用)
const vbaRunner = new VBARunner('source.bas', { useVirtualFS: false });
const fs = vbaRunner.evaluator.fs;
fs.writeFileSync('./sandbox/c/input.txt', "データ");
vbaRunner.run('ProcessFile', []);
const result = fs.readFileSync('./sandbox/c/output.txt', 'utf-8');
```

| モード | パス表記 | 説明 |
|--------|---------|------|
| **VFS** (`useVirtualFS: true`) | `/sandbox/c/input.txt` | メモリ内の仮想パス |
| **Node.js** (`useVirtualFS: false`) | `./sandbox/c/input.txt` | 実ファイルシステムの相対パス |

**詳細説明** は下の「[詳細: ファイルシステムモードの違い](#詳細-ファイルシステムモードの違い)」を参照してください。

## 詳細: ファイルシステムモードの違い

### VFS モード (useVirtualFS: true)

- **ストレージ**: メモリ上の仮想ファイルシステム (`MemoryFileSystem`)
- **パス形式**: 絶対パス `/sandbox/c/input.txt` を使用
- **特徴**:
  - テストが完全に分離され、他のテストやホスト OS に影響を与えない
  - パスは任意（`/foo`, `/bar` など）でも技術的には動作するが、VBA の `SandboxPath` が `C:\foo` → `/sandbox/c/foo` と変換するため、一貫性のため `/sandbox/c/` を使うことをお勧め

### Node.js モード (useVirtualFS: false)

- **ストレージ**: 実ファイルシステム (`NodeFileSystem`)
- **パス形式**: 相対パス `./sandbox/c/input.txt` を使用
- **特徴**:
  - `sandboxRoot` に指定したディレクトリをルートとする
  - 必ず **相対パス** を使用（絶対パス `/sandbox/c` は OS ルートを指すため危険）
  - `sandboxRoot` はデフォルト値 `/sandbox` だが、`new VBARunner(..., { sandboxRoot: './test-files' })` で上書き可能

### 実装の詳細

| 内部処理 | VFS | Node.js |
|---------|-----|---------|
| ファイルシステム | `MemoryFileSystem` | `NodeFileSystem` |
| `/sandbox/c/foo` の解釈 | メモリ上のキー | 実ディレクトリ `./sandbox/c/foo` |
| `SandboxPath.toRealPath('C:\\foo')` | `/sandbox/c/foo` (仮想) | `./sandbox/c/foo` (実) |
| 環境変数隔離 | ✅ 完全隔離 | ⚠️ `sandboxRoot` 内のみ |

## 高度な機能と仕様

本セクションでは、大規模なリファクタリングや複雑なテストケースで必要となる詳細な仕様について解説します。

### 複数モジュールの評価とスコープ

`VBARunner` に対して、ディレクトリを指定して複数のモジュールをまとめてロードすることができます。その際、以下の拡張子のファイルを対象とします。

| 拡張子 | 対象モジュール種別 |
|--------|------------------|
| `.bas` | 標準モジュール |
| `.cls` | クラスモジュール |
| `.frm` | フォームモジュール |

> **ファイルロード順**: アルファベット順で確定的にロードされます。モジュールレベルの初期化コード（`Sub`/`Function` 外）がある場合はこの順で実行されます。

### スコープ修飾子（Public / Private / Friend）について

`Public`・`Private`・`Friend` の修飾子はパースされ、プロシージャ宣言の AST に保持されます。

#### アクセス制御の仕様

| 呼び出し元 | Private プロシージャ | Public / 修飾子なし |
|-----------|-------------------|-------------------|
| `vbaRunner.run()` / `vbaRunner.eval()`（TypeScript） | ✅ 呼び出し可能 | ✅ 呼び出し可能 |
| **同一モジュール**内の VBA コード | ✅ 呼び出し可能 | ✅ 呼び出し可能 |
| **別モジュール**の VBA コード | ❌ 実行時エラー | ✅ 呼び出し可能 |

- **テストランナーからは `Private` でも直接呼び出せます**。ユニットテストの用途でプライベートなヘルパー関数も個別に検証できます。
- **VBA ソース同士のクロスモジュール呼び出しでは `Private` が制御されます**。別ファイルの `Private` プロシージャを呼んだ場合、以下のようなエラーが発生します：

  ```
  Execution error: Cannot call Private procedure 'PrivateHelper'
    from module 'ModuleB.bas' (defined in 'ModuleA.bas')
  ```

- `Friend` はパース・保持されますが、実行時のアクセス制御は `Public` と同等です（制御なし）。
- 修飾子なしのプロシージャは VBA 仕様どおり `Public` と同等です。

## VSCode拡張機能の使い方

### インストール・起動

1. このリポジトリを VSCode で開く
2. `F5` キーで「Extension Development Host」を起動する
3. 新しいウィンドウで `.bas` / `.cls` / `.frm` ファイルを開く

LSPの設計・実装方針については [`LSP.md`](LSP.md) を参照してください。

### 実装済みLSP機能

#### ホバー (Hover)

シンボル（Sub/Function/変数/定数/クラス/イベント）にマウスカーソルを合わせると、シグネチャをポップアップ表示します。

```vba
Sub CalcSum(a As Integer, b As Integer)
'         ↑ ここにホバー → "Sub CalcSum(a As Integer, b As Integer)" が表示される
```

#### 定義へ移動 (Go to Definition)

シンボルにカーソルを置いて `F12` を押すと、そのシンボルの宣言位置にジャンプします。呼び出し元から定義元への移動に使います。

| 対応シンボル | 例 |
|---|---|
| Sub / Function / Property | `Sub Foo()`, `Function Bar() As Long` |
| 変数宣言 | `Dim x As Integer`, `Public count As Long` |
| 定数 | `Const MAX As Long = 100` |
| クラス | `Class MyClass` |
| イベント | `Event DataChanged(newVal As Variant)` |

#### コード補完 (Completion)

入力中に VBA キーワード・組み込み関数・ソース内で定義されたプロシージャの候補を表示します。`.` を入力したときにもトリガーされます。

### 拡張機能のビルド

```bash
npm run build:extension   # dist/extension.cjs を生成
```

---

## UI画面の起動
ブラウザ上でVBAエディタ環境を立ち上げ、`Debug.Print` クラスの動作などを確認します。
これはおまけです。VBAを実行できることを示すデモのためにあります。

デプロイ済みのサイトは以下から利用できます:
👉 https://vba-web-runner.netlify.app/

ローカルで起動する場合:

```bash
# 依存関係のインストール
npm install

# 開発用ローカルサーバーの起動 (http://localhost:5173/)
npm run dev
```
---

## CLI(自動テスト)の実行

CLIからTypeScriptのテストランナーを利用し、VBAモジュールのユニットテストや言語仕様テストを実行します。

### 1. 個別のテスト実行

特定のテストファイル（例: サンプルのユニットテストや特定の仕様テスト）を実行する場合、`esbuild` でバンドルしてから `node` で実行します。

```bash
# サンプルのユニットテストを実行
npx esbuild sample/tests/ts/TaskScheduler_Core.test.ts --bundle --outfile=sample/tests/ts/TaskScheduler_Core.test.cjs --platform=node && node sample/tests/ts/TaskScheduler_Core.test.cjs

# 特定の仕様テスト（例: 文字列関数バリアント）を実行
npx esbuild tests/spec/mid-variants.test.ts --bundle --outfile=tests/spec/mid-variants.test.cjs --platform=node && node tests/spec/mid-variants.test.cjs
```

### 2. 全ての仕様テスト (spec) の一括実行

VBA Runner 自体の動作確認用として、`tests/spec/` ディレクトリには、VBAの言語仕様（文字列操作、数学関数、プロシージャ宣言等）を網羅するためのテストが多数配置されています。これらを一括で実行するためのスクリプトを用意しています。

```bash
# tests/spec/ 配下のすべてのテストを一括実行
./run_spec_tests.sh
```

実行後、すべてのテストで `✅ ... All tests passed!` と表示され、最後に `--- All spec tests completed successfully! ---` と出力されれば正常です。

# 補足: 技術的な詳細仕様

このセクションでは、VBA実行エンジンの内部動作や、詳細な仕様について解説します。

## VBA実行エンジンの動作仕様とインターフェース

本実行エンジンは、TypeScript/JavaScript環境からVBAコードを動的に評価・実行するための柔軟なインターフェースを提供しています。

### 1. VBAソースのコンパイルと実行 (`evaluate`)

VBAソースコード（`.bas`）をパースして得られたAST（抽象構文木）を順次実行します。

- **宣言の登録**: `Sub`, `Function`, `Property`, `Type`, `Const`, `Dim` などの宣言が出現すると、現在の実行環境（Environment）に登録されます。
- **トップレベル実行**: 素のVBA（Excel/Access等）では標準モジュールのトップレベルに実行文（宣言以外）を書くことはできませんが、本実行エンジンでは**トップレベルでの実行ステートメントを許可**しています。
  - Web UIの「Run」ボタンでは、ソースコードを上から順に実行するため、末尾に `MainSub` のように記述しておけば、コンパイルと同時にメインルーチンを開始できます。

### 2. `vbaRunner.run(procedureName, args)` の仕様

コンパイル済み環境に登録されているプロシージャを、JavaScript側から直接呼び出します。

- **引数の渡し方**: 第二引数にJavaScriptの配列を渡すと、VBA側のプロシージャ引数にマッピングされます。
- **戻り値**: `Function` または `Property Get` の戻り値をJavaScriptの値として返します。`Sub` の場合は `undefined` を返します。
- **用途**: 事前に定義したロジックを、異なるパラメータで繰り返しテストする場合に適しています。

### 3. `vbaRunner.eval(exprString)` の仕様

文字列として渡されたVBAコードを動的に評価します。

- **評価順序**:
  1. まず、入力を「VBA式（Expression）」としてパース・評価を試みます（例: `1 + 2`, `Join(arr, ",")`）。成功すればその値を返します。
  2. 式としてパースできない場合（例: `Dim` 宣言を含む、または複数行の処理）、自動的に「ステートメント（Statement）」としてパースし、`evaluate` を実行します。
- **特徴**: このフォールバック機構により、式だけでなく `vbaRunner.eval("x = 10 : debug.print x")` のようなアドホックなスクリプト実行も可能にしています。
## 制限事項

VBA Runnerは、VBAのロジックをリファクタリングし、モダンな言語環境でテスト可能にすることを目的としています。そのため、VBAの厳密な仕様とは以下の点で異なる恒久的な制限事項があります。

### 1. Boolean 型の扱い
VBAの `Boolean` 型は、内部的に `-1` (True) または `0` (False) を保持する特殊な型です。本エンジンではこれを `VbaBoolean` というラッパークラスで実装しています。

- **JSとの相互作用**: `vbaRunner.run()` 等の戻り値を直接 JavaScript の `if` 文で判定すると、オブジェクトであるため **`False` であっても `true` と評価されます**。
- **解決策**: JS側で真偽判定を行う場合は、`result.valueOf() !== 0` とするか、アサーションにおいて `assert.isTrue(result)` を使用してください。

### 2. 数値型（Integer / Long / Double）の未分化
VBAには `Integer` (16bit), `Long` (32bit), `Double` (64bit float) 等の厳密な数値型がありますが、本エンジンでは実行速度とシンプルさを優先し、**実行時の値はすべて JavaScript の `number` (64bit float) として扱います**。

- **型名の判定**: `TypeName` / `VarType` はリテラルや宣言型から型を推論します。
  - 数値リテラル: `TypeName(10)` → `"Integer"`、`TypeName(40000)` → `"Long"`、`TypeName(1.5)` → `"Double"`
  - 宣言された変数: `Dim x As Long : TypeName(x)` → `"Long"`
  - 式の結果: `TypeName(10 + 20)` → `"Double"`（演算結果はデフォルトで Double 扱い）
- **計算精度**: すべて Double 精度で計算されます。Integer / Long のオーバーフロー検出や型キャストによる切り捨ては行われません。

### 3. コレクション・配列のインデックス
- **配列**: VBA インデックスを JavaScript の配列インデックスにそのままマップします（`Option Base 1` や `Dim arr(L To U)` も含む）。JS 配列の `length` は常に `upper + 1` で、`[0]`〜`[lower-1]` は `undefined`、`[lower]`〜`[upper]` に値が入ります。詳細は [配列セクション](#配列) を参照してください。
- **コレクション**: `Scripting.Dictionary` は JS の `Map` をベースにしたプロキシオブジェクトとして動作します。

### 4. 外部コマンドの実行 (Shell)
- **制限**: セキュリティおよびユニットテストへの集中を目的とし、`Shell` 関数は外部コマンドを実際に実行しません。
- **挙動**: 実行しようとしたコマンドの内容を `Debug.Print` 相当の出力としてログに記録し、固定のタスクID（1）を返します。

## VBA でテストを書き、実環境で実行する

VBA で直接テストコードを記述し、実 VBA 環境（Excel など）で実行するための仕組みを提供しています。このアプローチは、**Excel マクロ環境で直接動作を検証したい場合**に適しています。

### VBA テストソースの構成ルール

`tests/spec/vba/` ディレクトリに VBA ファイルを配置します。ファイルの種類によって命名規則が異なります。

#### ファイル命名規則

| 種類 | 命名 | 例 |
|---|---|---|
| テストプロシージャを含む `.bas` | `<名前>Test.bas`（**Test サフィックス**） | `CurrencyOperationsTest.bas` |
| テスト専用サポートクラス `.cls` | 短い名前（`Test` サフィックスなし） | `RefA.cls`, `Helper.cls` |
| 共通インフラクラス `.cls` | 用途がわかる名前 | `AssertHelper.cls` |
| 生成されたランナー `.bas` | `<名前>Test_runner.bas` | `CurrencyOperationsTest_runner.bas` |

> **注意**: ファイル名から拡張子を除いた文字列が VBA のモジュール名になります。VBA のモジュール名は **31 文字以内** でなければなりません。

#### 複数ファイルが必要なテストはサブディレクトリに

1 つのテストで複数のクラスファイルが必要になる場合は、サブディレクトリを作成してファイル名を短くします。

```
tests/spec/vba/
├── AssertHelper.cls              ← 全テスト共通（assert オブジェクト）
├── CurrencyOperationsTest.bas    ← Currency 型テスト
├── CurrencyOperationsTest_runner.bas  ← 自動生成ランナー
├── SetupTeardownTest.bas
├── SetupTeardownTest_runner.bas
├── ArgCountTest.bas
├── ArgCountTest_runner.bas
└── Circular/                     ← 複数クラスが必要なテストはサブディレクトリ
    ├── RefA.cls
    ├── RefB.cls
    ├── Helper.cls
    ├── TerminateTest.bas
    └── TerminateTest_runner.bas
```

#### AssertHelper クラス（`tests/spec/vba/AssertHelper.cls`）

テストの検証には `AssertHelper` インスタンスを使用します。失敗時は `Debug.Print` でメッセージを出力してから `Err.Raise` を発生させます。

| メソッド | 用途 |
|---|---|
| `assert.Assert actual, expected, "メッセージ"` | 値比較 |
| `assert.IsTrue value, "メッセージ"` | 真であることの検証 |
| `assert.IsFalse value, "メッセージ"` | 偽であることの検証 |

### VBA テストコードの書き方

#### テストプロシージャの命名規則

- テストプロシージャは `Sub Test_<機能名>(assert)` の形式で定義
- プロシージャ名は `Test_` で始める（例: `Test_CurrencyArithmetic`）
- `assert` 引数に `AssertHelper` インスタンスが渡される

#### テストプロシージャの記述例

```vba
' tests/spec/vba/CurrencyOperationsTest.bas
Option Explicit

Sub Test_CurrencyArithmetic(assert)
    Dim c1 As Currency
    Dim c2 As Currency
    Dim result As Currency

    c1 = 100.25
    c2 = 50.75
    result = c1 + c2

    assert.Assert result, 151#, "100.25 + 50.75 = 151"
End Sub

Sub Test_CurrencyMultiplication(assert)
    Dim price As Currency
    Dim quantity As Currency
    Dim total As Currency

    price = 19.99
    quantity = 3
    total = price * quantity

    assert.Assert total, 59.97, "19.99 * 3 = 59.97"
End Sub

Sub Test_CurrencyDivision(assert)
    Dim total As Currency
    Dim count As Currency
    Dim average As Currency

    total = 100#
    count = 4
    average = total / count

    assert.Assert average, 25#, "100 / 4 = 25"
End Sub
```

#### SetUp と TearDown（テスト前後の処理）

テストランナーは **`SetUp`** と **`TearDown`** というサブルーチンを自動検出し、サポートします。

- **`SetUp` Sub**: 各テストの**実行前**に自動的に呼び出されます。テスト環境の初期化、テストデータの準備などに使用します。
- **`TearDown` Sub**: 各テストの**実行後**に自動的に呼び出されます。テスト後の後処理、リソース解放などに使用します。

```vba
' tests/spec/vba/SetupTeardownTest.bas
Option Explicit

Dim testCounter As Integer
Dim testState As String

Sub SetUp()
    testCounter = 0
    testState = "initialized"
End Sub

Sub TearDown()
    testCounter = 0
    testState = ""
End Sub

Sub Test_SetupInitializesCounter(assert)
    assert.Assert testCounter, 0, "SetUp counter=0"
End Sub

Sub Test_CounterIncrement(assert)
    testCounter = testCounter + 1
    testCounter = testCounter + 1
    assert.Assert testCounter, 2, "Counter increment"
End Sub
```

**実行フロー:**

```
SetUp()
  └─ Test_SetupInitializesCounter(assert) → PASS
TearDown()

SetUp()
  └─ Test_CounterIncrement(assert) → PASS
TearDown()
```

**ベストプラクティス:**

- SetUp で共通のテスト環境を準備し、各テストが独立して実行できるようにする
- TearDown でテスト実行時に作成されたリソースを確実に解放する
- SetUp/TearDown 内でエラーが発生した場合、`On Error Resume Next` で catch され、テストの実行は継続される
- `On Error Resume Next` が有効な間に `assert` メソッドを呼ぶと、`Err.Raise` が飲み込まれてテストが誤ってパスする。必ず `Err.Number` を変数に退避してから `On Error GoTo 0` を呼び、アサーションはその後に行うこと

### テストランナーの自動生成

`test-libs/vba-test-generator.ts` ツールを使用して、VBA テストファイル内のすべての `Test_*` プロシージャを検出し、**実 VBA 環境で動作するテストランナー Sub** を自動生成できます。生成されたランナーは `<名前>Test_runner.bas` として保存し、リポジトリにコミットします。

#### 1. 単一ファイルのテストランナーを生成

```bash
# esbuild でジェネレーターをビルドして実行（推奨）
./node_modules/.bin/esbuild test-libs/vba-test-generator.ts --bundle --outfile=test-libs/vba-test-generator.cjs --platform=node
node test-libs/vba-test-generator.cjs tests/spec/vba/CurrencyOperationsTest.bas tests/spec/vba/CurrencyOperationsTest_runner.bas

# モジュール修飾なし
node test-libs/vba-test-generator.cjs tests/spec/vba/CurrencyOperationsTest.bas tests/spec/vba/CurrencyOperationsTest_runner.bas --no-module-qualifier
```

**生成されるテストランナー（`CurrencyOperationsTest_runner.bas`）の例:**

```vba
Option Explicit

' Auto-generated test runner from vba-test-generator
' Run this Sub in Excel VBA environment to execute all tests
Sub RunAllTests()
    Dim allPass As Boolean
    Dim passCount As Integer
    Dim failCount As Integer
    Dim testResults As String
    Dim assert As New AssertHelper

    allPass = True
    passCount = 0
    failCount = 0
    testResults = "=== Test Results ===" & vbCrLf & vbCrLf

    ' Execute Test_CurrencyArithmetic
    On Error Resume Next
    Err.Clear
    CurrencyOperationsTest.Test_CurrencyArithmetic assert
    If Err.Number = 0 Then
        testResults = testResults & "[PASS] Test_CurrencyArithmetic" & vbCrLf
        passCount = passCount + 1
    Else
        testResults = testResults & "[FAIL] Test_CurrencyArithmetic" & vbCrLf
        failCount = failCount + 1
        allPass = False
    End If
    On Error GoTo 0

    ' ... (以降同様)

    Debug.Print testResults
    MsgBox testResults, IIf(allPass, vbInformation, vbCritical), "Test Results"
End Sub
```

**SetUp/TearDown がある場合:**

SetUp Sub と TearDown Sub を検出すると、生成されたテストランナーは各テストの前後にこれらを自動的に呼び出します：

```vba
    ' Execute Test_SetupInitializesCounter
    On Error Resume Next
    SetupTeardownTest.SetUp    ' ← 各テストの前に SetUp を呼び出し
    On Error GoTo 0
    On Error Resume Next
    Err.Clear
    SetupTeardownTest.Test_SetupInitializesCounter assert
    If Err.Number = 0 Then
        ...
    End If
    On Error GoTo 0
    On Error Resume Next
    SetupTeardownTest.TearDown  ' ← 各テストの後に TearDown を呼び出し
    On Error GoTo 0
```

#### 2. モジュール修飾（Module Qualification）

テストランナーは、デフォルトで**モジュール修飾**を使用してプロシージャを呼び出します。これにより、複数のモジュールで同じ名前のテスト関数が定義されている場合でも、それぞれを区別して実行できます。

```bash
# モジュール修飾あり（デフォルト）
node test-libs/vba-test-generator.cjs tests/spec/vba/CurrencyOperationsTest.bas

# モジュール修飾なし
node test-libs/vba-test-generator.cjs tests/spec/vba/CurrencyOperationsTest.bas --no-module-qualifier
```

#### 3. ディレクトリ内のすべてのテストファイルを処理

```bash
node test-libs/vba-test-generator.cjs --dir tests/spec/vba
```

`tests/spec/vba/` ディレクトリ内のすべての `.bas` ファイルに対してランナーを生成し、各ファイルの隣に `<filename>_runner.bas` として保存します。サブディレクトリは個別に実行してください：

```bash
node test-libs/vba-test-generator.cjs --dir tests/spec/vba/Circular
```

### 実 VBA 環境でテストを実行

生成されたテストランナー Sub を、Excel または他の VBA 環境で実行します。

#### 手順

1. **Excel を開き、VBA エディタを起動** （Alt + F11）
2. **テストファイルと生成されたテストランナーをインポート**
   - 生成された `*_runner.bas` ファイルをコピー
   - VBA エディタでモジュールを新規作成
   - テストランナーコードをペースト
3. **`RunAllTests` Sub を実行** （F5 キーまたは「実行」ボタン）
4. **結果を確認**
   - メッセージボックスに結果が表示されます
   - `Debug.Print` 出力は VBA エディタのイミディエイトウィンドウに表示されます

#### 実行例

```vba
' VBA Editor で RunAllTests Sub を実行すると以下のような結果が表示されます：
' === Test Results ===
'
' [PASS] TestCurrencyArithmetic
' [PASS] TestCurrencyMultiplication
' [PASS] TestCurrencyDivision
' [PASS] TestCurrencyTaxCalculation
' [PASS] TestCurrencyArray
' [PASS] TestCurrencyConversion
' [PASS] TestCurrencyComparison
' [PASS] TestCurrencyNegative
' [PASS] TestCurrencyPrecision
' [PASS] TestCurrencyFunctionParameter
'
' === Test Summary ===
' Total: 10
' Passed: 10
' Failed: 0
' === Test Complete ===
```

### テスト書き方のベストプラクティス

1. **テストプロシージャは単一責任**: 各テストプロシージャは 1 つの機能・シナリオのみをテストする
2. **テストデータは自動初期化**: テストプロシージャ内でテストデータを自動生成・初期化する（外部依存性を避ける）
3. **明示的な検証**: 処理結果を必ず `TestFunction = (条件)` で Boolean で返す
4. **独立実行可能**: テストプロシージャ間に依存関係を持たない（実行順序を変えても成功するように）
5. **エラーハンドリング**: 予期しないエラーが発生した場合、`On Error Resume Next` で catch し、テスト失敗として報告する

### ワークフロー例

```bash
# 1. テストファイルを作成
# tests/spec/vba/Test_MyFeature.bas

# 2. テストランナーを生成
npx ts-node test-libs/vba-test-generator.ts tests/spec/vba/Test_MyFeature.bas tests/spec/vba/Test_MyFeature_runner.bas

# 3. 生成されたテストランナーをテストコードと一緒に Excel に貼り込む
# (ファイルをコピーして VBA Editor にペーストするか、ファイルメニューでインポート)

# 4. Excel で RunAllTests を実行
# (VBA Editor の Run メニュー、または Alt+F5)

# 5. 結果をメッセージボックスと Debug.Print で確認
```

---

## さらなるドキュメント

詳細な設計ガイド・実践例・内部仕様は `docs/` 配下のドキュメントを参照してください。

### テスト設計の原則（基礎）

| ドキュメント | 内容 |
|------------|------|
| [TESTING_STRATEGY.md](docs/TESTING_STRATEGY.md) | VBA テスト設計の原則：Domain Logic と Excel I/O の分離、テスト境界の引き方 |
| [REFACTORING_GUIDE.md](docs/REFACTORING_GUIDE.md) | Excel 依存を含む VBA コードのリファクタリング手法（4 パターン） |

### 実装レベルのテスト（実践）

| ドキュメント | 内容 |
|------------|------|
| [TEST_FRAMEWORK_GUIDE.md](docs/TEST_FRAMEWORK_GUIDE.md) | TypeScript テストフレームワークを使った VBA テストの書き方 |
| [MOCK_GUIDE.md](docs/MOCK_GUIDE.md) | Excel オブジェクト（Sheets / Range / Application）のモック実装ガイド（3 段階：即使える / カスタマイズ / 設計論） |

### 具体例（学習用）

| ドキュメント | 内容 |
|------------|------|
| [REFACTORING_EXAMPLE.md](docs/REFACTORING_EXAMPLE.md) | TaskScheduler マクロを題材にしたリファクタリングの実例 |
| [INTEGRATION_TEST_EXAMPLE.md](docs/INTEGRATION_TEST_EXAMPLE.md) | TaskScheduler の統合テスト実装例 |

### 内部実装の詳細

| ドキュメント | 内容 |
|------------|------|
| [TYPE_SYSTEM_SPEC.md](docs/TYPE_SYSTEM_SPEC.md) | VBA 型システムの仕様と VBA Runner での実装方針 |
| [INDEX.md](docs/INDEX.md) | ドキュメント索引・用途別読み順ガイド |
