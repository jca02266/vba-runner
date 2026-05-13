# VBA Execution Environment

Excelがない環境で、作成したVBAコードの動作確認、リファクタリング、および自動テスト（ユニットテスト）を実行するためのプロジェクト（VBAパーサーおよびAST評価器）です。CLIによるTypeScriptのテストランナー環境を備えています。おまけでReactのWeb UIも用意しています。

## 本プロジェクトの目的
- **環境非依存での動作確認**: Excel (Windows/Mac) に依存せず、モダンなブラウザ上またはサーバー上(Node.js等)で直接VBAの構文とロジックを実行・検証できます。
- **リファクタリングの支援**: 巨大なVBAコードベースから純粋な関数・サブルーチンを安全に切り出し、モジュールを整理するための検証基盤を提供します。
- **ユニットテストの実行**: TypeScriptを利用したテストランナーを通じて、抽出されたVBA関数に対して直接モックデータやアサーションを評価し、プログラムによるテスト自動化を可能にします。

上記より、VS Codeなどのモダンな環境で（AIによる支援を受けながら）VBAソースのリファクタリングとテストの実行を可能にします。

### 今後の展望: VSCode統合開発環境 (Phase 2)
本プロジェクトにおける第一段階（MS-VBAL仕様書の100%実装）は完了しました。次なる最終目標として、このコンパイラエンジンとテスト基盤を組み込んだ **「VSCode完結のVBA統合開発環境（IDE）」** を提供する拡張機能の開発を構想しています。
LSP (Language Server Protocol) によるリアルタイムな構文チェックや補完、VSCode Testing APIへの統合に関するロードマップの詳細は [`TODO_NEXT.md`](TODO_NEXT.md) および [`LSP.md`](LSP.md) を参照してください。

## 基本方針：環境依存操作のスタブ・仮想化

本プロジェクトは**リファクタリングとユニットテストの実行**を主目的としています。そのため、実行環境（OSやファイルシステム）に副作用を及ぼす、または環境に依存する以下の操作は、原則として**サンドボックス化またはスタブ・フェイク実装**で行います。

- **外部コマンドの実行 (`Shell`)**: 実際のコマンドは実行せず、ログ出力と成功ステータスの返却のみを行います（スタブ化）。
- **ファイル操作 (`Open`, `Kill` 等)**: 指定されたサンドボックスディレクトリ (`sandbox/`) 内に制限されます（仮想化）。
- **環境変数 (`Environ`)**: OSの実環境変数にはアクセスせず、仮想的な環境変数（`.env` 等）を使用します。
- **レジストリ (`GetSetting`, `SaveSetting`)**: Windowsレジストリにはアクセスせず、メモリ上のマップ等による仮想レジストリを使用します。
- **外部オブジェクト・ネットワーク (`CreateObject`, `MSXML2.XMLHTTP`)**: 主要なライブラリをスタブ・フェイク実装として提供し、実際のCOMオブジェクト生成やネットワーク通信は行いません。
- **ユーザー対話 (`MsgBox`, `InputBox`)**: 実行をブロックせず、ログ出力とデフォルト値の返却のみを行います（スタブ化）。

AIを活用する場合、リファクタリングしたいソースコードと同じ階層に本プロジェクトを配置し、AIに対して「この `vba-compiler` を使ってリファクタリングおよびテストを実行して」と依頼（プロンプト）することで、スムーズに作業を進めることができます。

## ディレクトリ構成とVBAサンプルコード
本プロジェクトでは、巨大なVBAマクロを「テスト可能な単位」にリファクタリングする事例として、以下の構成を採用しています。

```
project-dir
 ├── vba-compiler          ...本プロジェクト(`git clone`で配置してください)
 │   ├── src/compiler/
 │   ├── sample/src/vba/
 │   ├── sample/src/vba_legacy/
 │   ├── sample/tests/ts/        
 │   ├── tests/spec/             
 │   ├── tests/test-libs-tests/  
 │   └── test-libs/              
 ├── node-vX.X.X-win-x64   ...実行用Node.jsの実行環境（Windows64bitスタンドアローン版の例）
 │   ├── node_modules/        普段Node.jsを使わない人はここにスタンドアローン版を配置するのが手軽です
 │   └── ...
 ├── massive-vba-project   ...リファクタリングしたいVBAソースのプロジェクト
 │   ├── Module1.vba
 │   ├── Module2.vba
 │   ├── ...
 │   └── Form1.frm
 └── test/                  ...VBAソースをテストするためのTypeScriptで書かれたテストプロジェクト
      ├── Module1Test/
      │   └── test.ts
      └── Module2Test/
          └── test.ts
```

本プロジェクトには、動作確認用のサンプルVBAソースとテストコードが以下のディレクトリに用意されています。これらを使用して、環境のセットアップが正しく完了しているかを確認できます。

- `src/compiler/` - TypeScriptで書かれたVBA用Lexer、Parser、Evaluatorのコアエンジン群。
- `sample/src/vba/` - 本番稼働を想定したプロダクションコード。
  - `TaskScheduler.vba`: 処理の見通しが良くなったリファクタリング**後**のメインルーチン。
  - `TaskScheduler_Core.vba`: `TaskScheduler.vba` から抽出された、純粋なビジネスロジック関数モジュール（**ユニットテストの対象**）。
- `sample/src/vba_legacy/` - 参照用のレガシーコード。
  - `TaskScheduler_v1.vba`: リファクタリング**前**の、機能が密結合した巨大なオリジナルソース。
- `sample/tests/ts/` - Node.js 上でVBAスクリプトをAST評価し、自動テストを走らせるTypeScriptスクリプト群。
  - `TaskScheduler_Core.test.ts`: `TaskScheduler_Core.vba` 内の各関数の振る舞いを検証するユニットテスト。
- `test-libs/` - 汎用のテストランナーユーティリティ（VBATest クラス、assert 等）。

また、以下にVBA評価器自体のテストが配置されています。

- `tests/spec/` - VBAの仕様を確認するための各機能のテスト。
- `tests/test-libs-tests/` - `test-libs/` 配下のテストランナー自身の機能テスト。

## テストの書き方

VBAソースのテストを記述する際は、テストランナー `vba-compiler/test-libs/test-runner.ts` を使用します。

> [!TIP]
> 実際にコードを書く前に、ブラウザ上で動作を確認したい場合は [Web UI デモサイト](https://vba-web-compiler.netlify.app/) をご利用ください。`Debug.Print` の結果や構文チェックを即座に試すことができます。

### 1. `eval`: VBAの式やコード断片をその場で評価する

VBAの構文をそのまま文字列として渡し、評価結果を取得します。最も手軽にVBAエンジンを試すことができる方法です。

```typescript
import { VBATest } from '../../test-libs/test-runner';
const vbaTest = new VBATest(''); // 空の環境を作成

// VBAの計算式を直接評価
const sum = vbaTest.eval("1 + 2 + 3"); // => 6

// 変数宣言や代入を含む、複数行の処理も実行可能
vbaTest.eval("Dim x : x = 10 : Debug.Print x * 2");
```

### 2. `run`: ロードしたソース内の関数を実行する

既存のVBAファイル（`.vba`）をロードし、そこに定義されたプロシージャを引数を指定して呼び出します。複雑なビジネスロジックのユニットテストに最適です。

```typescript
// 1. テスト対象のVBAファイルをロード
const vbaTest = new VBATest('sample/src/vba/Sample.vba');

// 2. 関数名を指定して実行（第2引数はJavaScriptの配列として引数を渡す）
const result1 = vbaTest.run('Add', [1, 2]);
const result2 = vbaTest.run('Multiply', [result1, 2]);

// 3. 結果をアサート
assert.strictEqual(result1, 3);
assert.strictEqual(result2, 6);
```

### 3. 複数ソースの一括ロード

既存のVBAファイル（`.vba`）を格納したディレクトリを指定すれば、配下のソースファイルをすべて読み込みます。
規模の大きなVBAプロジェクトではこの使い方になります。

```typescript
// 1. テスト対象のVBAファイルをロード
const vbaTest = new VBATest('sample/src/vba');

// 2. 関数名を指定して実行（第2引数はJavaScriptの配列として引数を渡す）
const result1 = vbaTest.run('Add', [1, 2]);
const result2 = vbaTest.run('Multiply', [result1, 2]);

// 3. 結果をアサート
assert.strictEqual(result1, 3);
assert.strictEqual(result2, 6);
```

## テスト環境での `eval` と `run` の使い分け
TypeScript側からVBAのロジックを呼び出す際、用途に応じて以下の2つのメソッドを利用します。

- **`vbaTest.eval(expressionString)`**:
  - **用途**: VBAの「式（Expression）」や「文（Statement）」を文字列として渡し、その場でVBAとして評価（パース）して実行します。
  - **特徴**: 
    - `AddNumbers(3, 4)` のように、VBAとして有効な式を評価し、その戻り値をダイレクトに取得する場合に向いています。
    - また、`MainLoop` のように括弧なしで「文」として実行されるVBA特有の呼び出しも自動で判別して実行します（この場合の戻り値は `undefined` となります）。
  - **例**: `const result = vbaTest.eval("AddNumbers(3, 4)");` または `vbaTest.eval("MainLoop");`

- **`vbaTest.run(procedureName, args)`**:
  - **用途**: VBAのサブルーチン（`Sub`）や関数（`Function`）に、TypeScript側からの配列引数を渡して実行します。
  - **特徴**: トップレベルから直接プロシージャ名を指定して呼び出すための「標準的」な実行メソッドです。
  - **例**: `const result = vbaTest.run("AddNumbers", [100, 200]);`

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
- **`eval` の強み**: `vbaTest.eval("Add(1, 2) * 10")` のように、戻り値を利用した計算式や複数行の処理を、VBA の記法でそのまま記述・評価できます。

> **Tip**: 基本的には、引数があるプロシージャのテストには `run` を使用します。サブルーチン（`Sub`）を実行した後に、その影響で書き換わった**グローバル変数やモジュール変数の値を個別にアサート（確認）したい場合**などは、`eval("VariableName")` を組み合わせて使うのが定石です。
> また、長期的には `run` だけで完結できるような「副作用のない純粋な関数（Function）」にロジックを切り出し、リファクタリングしていくことが推奨されます。

## テストコードで使う固有型と値

VBAの型システムはJavaScriptと異なるため、`run` / `eval` の戻り値をアサートする際に以下の対応表を参照してください。

### スカラー値の対応

| VBA の値          | JavaScriptの型・値    | 説明                                                       |
|-------------------|-----------------------|------------------------------------------------------------|
| `Empty`           | `null`（`vbaEmpty`）  | 未初期化のスカラー変数。`evaluator.ts`の`vbaEmpty`定数と同値 |
| `Nothing`         | `null`                | 未設定のオブジェクト参照                                   |
| `True`            | `VbaBoolean`オブジェクト | 詳細は[制限事項](#制限事項)を参照                         |
| `False`           | `VbaBoolean`オブジェクト | 詳細は[制限事項](#制限事項)を参照                         |
| 数値              | `number`              | 詳細は[制限事項](#制限事項)を参照                         |
| 文字列            | `string`              | ダブルクォート除去済みの生文字列                           |
| `Null`            | `Symbol`（`vbaNull`） | `vbaNull`定数と同値                                       |

```typescript
import { vbaEmpty } from '../../src/compiler/evaluator';

assert.strictEqual(result, vbaEmpty);   // VBAのEmpty（null）と比較
assert.strictEqual(result, null);       // 同上（直接nullでも可）
```

### ブール値のアサーション

VBAの比較演算や `IsNull`、`CBool` などの関数は `-1`（True）または `0`（False）を返します。
`assert.isTrue` / `assert.isFalse` を使うと、`vbaTrue`/`vbaFalse` の定数をインポートせずに簡潔にアサートできます。

```typescript
import { assert } from '../../test-libs/test-runner';

assert.isTrue(vbaTest.eval('1 = 1'),        '等値比較');
assert.isFalse(vbaTest.eval('IsNull(123)'), 'IsNull(123)');

// assert.strictEqual で書く場合は vbaTrue/vbaFalse をインポートする必要がある
// import { vbaTrue, vbaFalse } from '../../src/compiler/evaluator';
// assert.strictEqual(result, vbaTrue);
```

### 配列

本コンパイラの配列表現の基本ルール：**VBA インデックスと JavaScript の配列インデックスを一致させる**。
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
const arr = vbaTest.run('GetArray', []) as any[];
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
const arr = vbaTest.run('GetArray', []) as any[];
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
const arr = vbaTest.run('GetArr1To3', []) as any[];
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
const dict = vbaTest.run('BuildDict', []) as any;
assert.strictEqual(dict.__map__.get('key1'), 'value1');
assert.isTrue(dict.exists('key1'));
```

### CreateObject のファクトリ差し替え（モック登録）

VBA の `CreateObject(progId)` が返すオブジェクトを **テストコード側でモックに差し替え** できます。これにより、本体（`src/compiler/evaluator.ts`）に手を入れずに、テスト固有の COM オブジェクトをエミュレートできます。

#### API
- `Evaluator.registerExternalObject(progId, factory)`
- `VBATest.registerExternalObject(progId, factory)` （上記のプロキシ）

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
import { vbaTrue, vbaFalse } from '../src/compiler/evaluator';

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
import { VBATest, assert } from '../../test-libs/test-runner';
import { createRegExpMock } from '../../test-libs/regexp-mock';

const vbaTest = new VBATest('macro.vba');
// 1 回の登録で「CreateObject」も「New / Dim As」も同じ factory が使われる
vbaTest.registerExternalObject('VBScript.RegExp', createRegExpMock);

const result = vbaTest.run('CountDigits', ['abc 123 def 456']);
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

> **動作確認**: `tests/test-libs-tests/regexp-mock.test.ts` に `CreateObject` 形式と `New RegExp` 形式の両方を網羅したテストがあります。
> 同じ仕組みで `MSXML2.XMLHTTP` や自社の独自 COM オブジェクトもモック化できます。

---

## VBA 実行時の終了ステータスと例外の検知
 
TypeScript 側から `vbaTest.run()` または `vbaTest.eval()` を呼び出した際、VBA の終了状態（正常終了、`End` による強制終了、実行時エラー）を以下のように区別して検知できます。
 
### 1. 正常終了（Sub/Function の戻り値）
 
プロシージャが最後まで実行された、または `Exit Sub/Function` で抜けた場合、戻り値が TypeScript 側に返ります。
 
```typescript
try {
    const result = vbaTest.run("MyFunction", [10]);
    console.log("正常終了:", result);
} catch (e) {
    // 正常終了時はここには来ない
}
```
 
### 2. `End` ステートメントによる強制終了
 
VBA コード内で `End` ステートメントが実行されると、呼び出し側に `{ type: 'Terminate' }` というオブジェクトがスローされます。これは通常のランタイムエラーとは区別してハンドルできます。
 
```typescript
try {
    vbaTest.eval("End");
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
    vbaTest.run("NonExistentSub", []);
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
assert.strictEqual(vbaTest.run('GetX', []), 0);
```

## ファイル入出力のSandbox方針

VBAコードからのファイル入出力はすべて**Sandboxディレクトリ**内に限定します。これにより、実行環境のファイルシステムを保護しつつ、VBAのファイル操作命令をテスト可能にします。

### Sandboxの動作ルール

1. **Sandboxルートの指定**: Evaluatorの初期化時にSandboxルートディレクトリを指定します。デフォルトは `sandbox/`（プロジェクトルートからの相対パス）です。

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

本プロジェクトでは、ブラウザ上での動作や、テストの安全性を高めるために **仮想ファイルシステム (VFS)** を導入しています。

### 将来構想：VFS を基本とした実行環境
長期的には、Node.js 環境においても実際のディスクを操作する `NodeFileSystem` ではなく、メモリ上で完結する `MemoryFileSystem` (VFS) をデフォルトの実行環境にすることを構想しています。これにより、以下のメリットが得られます：
- **安全性**: テスト実行によってホスト OS のファイルが誤って削除・上書きされるリスクをゼロにします。
- **再現性**: テスト実行ごとにクリーンなファイルシステム状態から開始でき、並列実行も容易になります。
- **ブラウザ互換性**: テストコードそのものをブラウザ上でそのまま実行可能になります。

### VFS の有効化方法
`VBATest` の初期化時に `useVirtualFS: true` を指定するか、環境変数 `USE_VFS=1` を設定することで VFS モードになります。

```typescript
// 個別にVFSを有効化
const vbaTest = new VBATest('source.vba', { useVirtualFS: true });

// または環境変数で一括切り替え
// USE_VFS=1 node tests/spec/xxx.test.cjs
```

### JavaScript から VFS への直接アクセス (テストデータの準備)

テストの準備や実行結果の検証のために、JavaScript から VFS に対して直接ファイルを読み書きできます。

**簡潔な例**（VFS と Node.js 両対応）:

```typescript
import { VBATest } from './test-libs/test-runner';

// VFS モード
const vbaTest = new VBATest('source.vba', { useVirtualFS: true });
const fs = vbaTest.evaluator.fs;
fs.writeFileSync('/sandbox/c/input.txt', "データ");
vbaTest.run('ProcessFile', []);
const result = fs.readFileSync('/sandbox/c/output.txt', 'utf-8');
```

```typescript
// Node.js モード (実ファイルを使用)
const vbaTest = new VBATest('source.vba', { useVirtualFS: false });
const fs = vbaTest.evaluator.fs;
fs.writeFileSync('./sandbox/c/input.txt', "データ");
vbaTest.run('ProcessFile', []);
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
  - `sandboxRoot` はデフォルト値 `/sandbox` だが、`new VBATest(..., { sandboxRoot: './test-files' })` で上書き可能

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

`VBATest` に対して、ディレクトリを指定して複数のモジュールをまとめてロードすることができます。その際、以下の拡張子のファイルを対象とします。

| 拡張子 | 対象モジュール種別 |
|--------|------------------|
| `.vba` | 標準モジュール |
| `.cls` | クラスモジュール |
| `.frm` | フォームモジュール |

> **ファイルロード順**: アルファベット順で確定的にロードされます。モジュールレベルの初期化コード（`Sub`/`Function` 外）がある場合はこの順で実行されます。

### スコープ修飾子（Public / Private / Friend）について

`Public`・`Private`・`Friend` の修飾子はパースされ、プロシージャ宣言の AST に保持されます。

#### アクセス制御の仕様

| 呼び出し元 | Private プロシージャ | Public / 修飾子なし |
|-----------|-------------------|-------------------|
| `vbaTest.run()` / `vbaTest.eval()`（TypeScript） | ✅ 呼び出し可能 | ✅ 呼び出し可能 |
| **同一モジュール**内の VBA コード | ✅ 呼び出し可能 | ✅ 呼び出し可能 |
| **別モジュール**の VBA コード | ❌ 実行時エラー | ✅ 呼び出し可能 |

- **テストランナーからは `Private` でも直接呼び出せます**。ユニットテストの用途でプライベートなヘルパー関数も個別に検証できます。
- **VBA ソース同士のクロスモジュール呼び出しでは `Private` が制御されます**。別ファイルの `Private` プロシージャを呼んだ場合、以下のようなエラーが発生します：

  ```
  Execution error: Cannot call Private procedure 'PrivateHelper'
    from module 'ModuleB.vba' (defined in 'ModuleA.vba')
  ```

- `Friend` はパース・保持されますが、実行時のアクセス制御は `Public` と同等です（制御なし）。
- 修飾子なしのプロシージャは VBA 仕様どおり `Public` と同等です。

## UI画面の起動
ブラウザ上でVBAエディタ環境を立ち上げ、`Debug.Print` クラスの動作などを確認します。
これはおまけです。VBAを実行できることを示すデモのためにあります。

デプロイ済みのサイトは以下から利用できます:
👉 https://vba-web-compiler.netlify.app/

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

本プロジェクト(VBAコンパイラ)自体の動作確認用として、`tests/spec/` ディレクトリには、VBAの言語仕様（文字列操作、数学関数、プロシージャ宣言等）を網羅するためのテストが多数配置されています。これらを一括で実行するためのスクリプトを用意しています。

```bash
# tests/spec/ 配下のすべてのテストを一括実行
./run_spec_tests.sh
```

実行後、すべてのテストで `✅ ... All tests passed!` と表示され、最後に `--- All spec tests completed successfully! ---` と出力されれば正常です。

# 補足: 技術的な詳細仕様

このセクションでは、VBA実行エンジンの内部動作や、詳細な仕様について解説します。

## VBA実行エンジンの動作仕様とインターフェース

本コンパイラは、TypeScript/JavaScript環境からVBAコードを動的に評価・実行するための柔軟なインターフェースを提供しています。

### 1. VBAソースのコンパイルと実行 (`evaluate`)

VBAソースコード（`.vba`）をパースして得られたAST（抽象構文木）を順次実行します。

- **宣言の登録**: `Sub`, `Function`, `Property`, `Type`, `Const`, `Dim` などの宣言が出現すると、現在の実行環境（Environment）に登録されます。
- **トップレベル実行**: 素のVBA（Excel/Access等）では標準モジュールのトップレベルに実行文（宣言以外）を書くことはできませんが、本コンパイラでは**トップレベルでの実行ステートメントを許可**しています。
  - Web UIの「Run」ボタンでは、ソースコードを上から順に実行するため、末尾に `MainSub` のように記述しておけば、コンパイルと同時にメインルーチンを開始できます。

### 2. `vbaTest.run(procedureName, args)` の仕様

コンパイル済み環境に登録されているプロシージャを、JavaScript側から直接呼び出します。

- **引数の渡し方**: 第二引数にJavaScriptの配列を渡すと、VBA側のプロシージャ引数にマッピングされます。
- **戻り値**: `Function` または `Property Get` の戻り値をJavaScriptの値として返します。`Sub` の場合は `undefined` を返します。
- **用途**: 事前に定義したロジックを、異なるパラメータで繰り返しテストする場合に適しています。

### 3. `vbaTest.eval(exprString)` の仕様

文字列として渡されたVBAコードを動的に評価します。

- **評価順序**:
  1. まず、入力を「VBA式（Expression）」としてパース・評価を試みます（例: `1 + 2`, `Join(arr, ",")`）。成功すればその値を返します。
  2. 式としてパースできない場合（例: `Dim` 宣言を含む、または複数行の処理）、自動的に「ステートメント（Statement）」としてパースし、`evaluate` を実行します。
- **特徴**: このフォールバック機構により、式だけでなく `vbaTest.eval("x = 10 : debug.print x")` のようなアドホックなスクリプト実行も可能にしています。
## 制限事項 (Limitations)

本プロジェクトは、VBAのロジックをリファクタリングし、モダンな言語環境でテスト可能にすることを目的としています。そのため、VBAの厳密な仕様とは以下の点で異なる恒久的な制限事項があります。

### 1. Boolean 型の扱い
VBAの `Boolean` 型は、内部的に `-1` (True) または `0` (False) を保持する特殊な型です。本エンジンではこれを `VbaBoolean` というラッパークラスで実装しています。

- **JSとの相互作用**: `vbaTest.run()` 等の戻り値を直接 JavaScript の `if` 文で判定すると、オブジェクトであるため **`False` であっても `true` と評価されます**。
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
