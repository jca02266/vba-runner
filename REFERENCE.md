# REFERENCE — VBA Runner 詳細仕様

詳細な仕様・設計方針・高度な使い方をまとめたリファレンスです。
概要は [README.md](README.md) を参照してください。

---

## 基本方針：環境依存操作のスタブ・仮想化

VBA Runnerは**リファクタリングとユニットテストの実行**を主目的としています。そのため、実行環境（OSやファイルシステム）に副作用を及ぼす、または環境に依存する以下の操作は、原則として**サンドボックス化またはスタブ・フェイク実装**で行います。

- **外部コマンドの実行 (`Shell`)**: 実際のコマンドは実行せず、ログ出力と成功ステータスの返却のみを行います（スタブ化）。
- **ファイル操作 (`Open`, `Kill` 等)**: 指定されたサンドボックスディレクトリ (`sandbox/`) 内に制限されます（仮想化）。
- **環境変数 (`Environ`)**: OSの実環境変数にはアクセスせず、仮想的な環境変数（`.env` 等）を使用します。
- **レジストリ (`GetSetting`, `SaveSetting`)**: Windowsレジストリにはアクセスせず、メモリ上のマップ等による仮想レジストリを使用します。
- **外部オブジェクト・ネットワーク (`CreateObject`, `MSXML2.XMLHTTP`)**: 主要なライブラリをスタブ・フェイク実装として提供し、実際のCOMオブジェクト生成やネットワーク通信は行いません。
- **ユーザー対話 (`MsgBox`, `InputBox`)**: 実行をブロックせず、ログ出力とデフォルト値の返却のみを行います（スタブ化）。

AIを活用する場合、リファクタリングしたいソースコードと同じ階層に VBA Runner を配置し、AIに対して「この `vba-runner` を使ってリファクタリングおよびテストを実行して」と依頼（プロンプト）することで、スムーズに作業を進めることができます。

---

## VBA言語仕様の拡張機能

### クラス定義の拡張構文

実VBAの標準仕様では、クラスはファイル名（`.cls` ファイル）またはVB_Name属性によってのみ定義されます。`Class...End Class` という構文は存在しません。

本実行エンジンの**拡張機能**として、`.bas` モジュール内に `Class` キーワードを使用してクラスを定義可能にしています。これにより、テストコード内で複数のクラス定義を1つのファイルにまとめることができ、ユニットテスト環境での利便性が向上します。

```vb
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

---

## VBA仕様上の制限

本エンジンで実装されているVBAは、[MS-VBAL（MS-VBA Language）仕様 rev.2.4](https://learn.microsoft.com/en-us/openspecs/microsoft_general_purpose_programming_languages/ms-vbal/d5418146-0bd2-45eb-9c7a-fd9502722c74) に準拠しています。

### モジュール名の長さ制限

VBA仕様により、モジュール名（`.cls` ファイルの場合はファイル名、`.bas` ファイルの場合は名前属性）の最大長は **31文字** に制限されています。

```vb
' OK: 31文字以下
Module MyModule                    ' 8文字
Class VeryLongModuleNameExample    ' 27文字

' NG: 32文字以上
' Class ThisIsAVeryLongModuleNameThatExceeds31Chars
```

---

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

### Sub と Function の戻り値の対応表

| 呼び出し対象 | `run` の書き方（例） | `eval` の書き方（例） | 戻り値 |
| :--- | :--- | :--- | :--- |
| **Sub** | `run("Log", [10, "msg"])` | `eval("Log 10, ""msg""")` | `undefined` |
| **Function** | `run("Add", [1, 2])` | `eval("Add(1, 2)")` | 戻り値 |
| **VBA式** | (非対応) | `eval("1 + 2 + Add(1, 2)")` | 評価・計算結果 |
| **VBA文** | (非対応) | `eval("Dim x : x = 10 : Debug.Print x")` | `undefined` |
| **変数評価** | (非対応) | `eval("ActiveSheetName")` | 変数の現在の値 |

- **`run` の強み**: TypeScript 側の変数（配列やオブジェクト、数値など）を、そのまま第 2 引数の配列として安全に渡せます。
- **`eval` の強み**: `vbaRunner.eval("Add(1, 2) * 10")` のように、戻り値を利用した計算式や複数行の処理を、VBA の記法でそのまま記述・評価できます。

> **Tip**: 引数があるプロシージャのテストには `run` を使用します。サブルーチン（`Sub`）を実行した後に、その影響で書き換わった**グローバル変数やモジュール変数の値を個別にアサート（確認）したい場合**は、`eval("VariableName")` を組み合わせて使うのが定石です。

---

## テストコードで使う固有型と値

VBAの型システムはJavaScriptと異なるため、`run` / `eval` の戻り値をアサートする際に以下の対応表を参照してください。

### スカラー値の対応

| VBA の値 | JavaScriptの型・値 | 説明 |
|---|---|---|
| `Empty` | `null`（`vbaEmpty`） | 未初期化のVariant型変数。`evaluator.ts`の`vbaEmpty`定数と同値 |
| `Nothing` | `null` | 未設定のオブジェクト参照 |
| `True` | `VbaBoolean`オブジェクト | 詳細は[制限事項](#制限事項)を参照 |
| `False` | `VbaBoolean`オブジェクト | 詳細は[制限事項](#制限事項)を参照 |
| 数値 | `number` | 詳細は[制限事項](#制限事項)を参照 |
| 文字列 | `string` | ダブルクォート除去済みの生文字列 |
| `Null` | `Symbol`（`vbaNull`） | `vbaNull`定数と同値 |

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
assert.isTrue(vbaTrue);                        // ✅ OK
assert.isTrue(vbaRunner.eval('1 = 1'));        // ✅ OK（比較演算は vbaTrue を返す）
assert.isTrue(-1);                             // ❌ FAIL（数値 -1 は vbaTrue シングルトンとは別物）
assert.isTrue(vbaRunner.eval('Not 5'));        // ❌ FAIL（Not <数値> はビット反転の数値 -6 を返す）
```

「数値が **真とみなせる**（truthy）こと」を確認したいときは `assert.ok` を使います。

```typescript
assert.ok(vbaRunner.eval('5'));       // ✅ OK（5 は truthy）
assert.ok(vbaRunner.eval('Not 5'));   // ✅ OK（-6 は truthy）
assert.ok(vbaRunner.eval('0'));       // ❌ FAIL（0 は falsy）
```

#### VBA の Truthy/Boolean 仕様（補足）

| 文脈 | 例 | 結果 |
|------|-----|------|
| **代入時の暗黙型変換** | `Dim b As Boolean : b = 5` | `b = -1`（非 0 → True に自動変換） |
| **`If` 条件の Boolean 評価** | `If 5 Then ...` | 実行される（非 0 は True 扱い） |
| **比較演算 `= True` / `= False`** | `5 = True` | **False**（数値 5 と -1 は等しくない） |

### `Not` 演算子の戻り値

`Not` は引数の型によって戻り値の型が変わります（[MS-VBAL](https://learn.microsoft.com/en-us/openspecs/microsoft_general_purpose_programming_languages/ms-vbal/d5418146-0bd2-45eb-9c7a-fd9502722c74) §5.6.9.8.1）:

- **Boolean オペランド** → Boolean 反転（シングルトン）: `Not vbaTrue` → `vbaFalse`
- **数値オペランド** → ビット反転（数値）: `Not 5` → `-6`、`Not 0` → `-1`（数値の `-1` であって Boolean シングルトンではない）

`Not <数値>` の結果を **Boolean として** assert したい場合は、VBA 側で `CBool(...)` を挟むか、TypeScript 側で `assert.ok(...)` を使ってください。

### 配列

本実行エンジンの配列表現の基本ルール：**VBA インデックスと JavaScript の配列インデックスを一致させる**。
JavaScript 配列の `length` は常に `upper + 1`、`[0]`〜`[lower-1]` は未使用（`undefined`）、`[lower]`〜`[upper]` に値が入ります。

#### デフォルト (`Option Base 0`) — `Dim arr(N)`

`Dim arr(N)` は `arr(0)`〜`arr(N)` の `N + 1` 要素配列。

```typescript
const arr = vbaRunner.run('GetArray', []) as any[];
assert.strictEqual(arr.length, 6);    // N=5 のため要素数は 6
assert.strictEqual(arr[1], "first");
```

#### `Option Base 1` — `Dim arr(N)`

モジュール先頭で `Option Base 1` を宣言すると、JS 配列も `length = N + 1` のまま `[0]` は未使用、`[1]`〜`[N]` に値が入ります。

```typescript
const arr = vbaRunner.run('GetArray', []) as any[];
assert.strictEqual(arr.length, 4);
assert.strictEqual(arr[0], undefined);  // Option Base 1 のため [0] は未使用
assert.strictEqual(arr[1], "a");
```

#### `Dim arr(L To U)` — 任意の下限

`length = U + 1`、`[0]`〜`[L-1]` は未使用、`[L]`〜`[U]` が有効。

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

```typescript
const dict = vbaRunner.run('BuildDict', []) as any;
assert.strictEqual(dict.__map__.get('key1'), 'value1');
assert.isTrue(dict.exists('key1'));
```

---

## 外部ライブラリの参照設定と `New ProjectName.ClassName`

### VBA における参照設定とは

VBA IDE の **[ツール] → [参照設定]** でライブラリを追加すると、オブジェクト生成に2通りの構文が使えます：

```vba
' 方法1: CreateObject — 遅延バインディング（参照設定不要）
Dim d As Object
Set d = CreateObject("Scripting.Dictionary")

' 方法2: New 修飾名 / Dim As 型名 — 早期バインディング（参照設定が必要）
Dim d As New Scripting.Dictionary   ' Dim + New
Set d = New Scripting.Dictionary    ' Set + New
Dim d As Scripting.Dictionary       ' 型名だけ（= Nothing で初期化）
```

本エンジンでは `registerComObject(factory)` でファクトリを登録することで、両方の構文が動作します。

### プロジェクト名前空間の自動登録

`registerComObject(() => ({ __progId__: 'Scripting.Dictionary', ... }))` のように `ProjectName.ClassName` 形式の `__progId__` を持つオブジェクトを登録すると：

- **プロジェクト名（`Scripting`）が自動的に「ライブラリプロジェクト」として登録される**
- `New Scripting.Dictionary` → factory が呼ばれてインスタンスを生成
- `CreateObject("Scripting.Dictionary")` → 同じ factory を使用
- `New Dictionary`（短縮形）→ `__className__: 'Dictionary'` が登録されていれば使用可
- `VarType(Scripting)` → **コンパイルエラー**「プロジェクトではなく、変数またはプロシージャを指定してください」（VBA 仕様通り）
- `Scripting.Dictionary` の修飾アクセス形式は引き続き正常動作

```typescript
// カスタムライブラリの登録例
vbaRunner.registerComObject(() => ({
    __progId__: 'MyLib.MyClass',  // CreateObject("MyLib.MyClass") の照合キー
    __className__: 'MyClass',     // New MyClass / Dim x As MyClass でも使える
    doSomething: () => 42,
}));
// → New MyLib.MyClass / CreateObject("MyLib.MyClass") が使えるようになる
// → VarType(MyLib) はエラーになる（VBA 仕様通り）
```

### MS-VBAL 上のプロジェクト分類

VBA の名前空間として機能する「プロジェクト」は MS-VBAL で3種類定義されています：

| 種別 | 例 | 説明 |
|---|---|---|
| **ライブラリプロジェクト** | `VBA`、`Scripting`、`ADODB` | 参照設定した型ライブラリ。`VBA` は予約名 |
| **ホストプロジェクト** | `Excel`、`Word` | ホストアプリのオブジェクトモデル |
| **ソースプロジェクト** | `VBAProject`（ユーザー命名可） | ユーザーの VBA コード本体 |

当エンジンはホストプロジェクト（`Excel` 等）を実装しないため、`registerComObject` で明示登録したライブラリのみがライブラリプロジェクトとして認識されます。`VBA` は起動時に常に登録済みです。

### 組み込みで使用可能なライブラリ

以下は `registerComObject` を呼ばずに最初から動作します：

| ProgID | `New` 修飾形式 | 短縮 `New` | 備考 |
|---|---|---|---|
| `Scripting.Dictionary` | `New Scripting.Dictionary` | `New Dictionary` | Map ベースの簡易実装 |
| `Scripting.FileSystemObject` | `New Scripting.FileSystemObject` | `New FileSystemObject` | Sandbox VFS に委譲 |
| `MSXML2.XMLHTTP` | `New MSXML2.XMLHTTP` | `New XMLHTTP` | スタブ（実通信なし） |
| `Microsoft.XMLHTTP` | `New Microsoft.XMLHTTP` | — | 同上 |
| `ADODB.Stream` | `New ADODB.Stream` | `New Stream` | スタブ |

テストでこれらをモックに差し替えたい場合は、同じ progId を持つファクトリで `registerComObject` を呼ぶと **組み込みの実装を上書き** できます。

---

## CreateObject のファクトリ差し替え（モック登録）

VBA の `CreateObject(progId)` が返すオブジェクトを **テストコード側でモックに差し替え** できます。

### API

- `Evaluator.registerComObject(factory, ...extraProgIds)`
- `VBARunner.registerComObject(factory)` （上記のプロキシ）

`factory` は `CreateObject` が呼ばれるたびに **新しいオブジェクト** を返す関数です。返すオブジェクトには `__progId__: string` が必須で、これが `CreateObject(progId)` の照合キーになります。登録されたファクトリは既存の組み込みスタブ（`Scripting.Dictionary` 等）よりも優先されます。

### モック実装の作り方

- **プロパティ**: 通常の JS オブジェクトのフィールドとして書きます。Boolean は `vbaTrue / vbaFalse`（`-1 / 0`）を使う。
- **メソッド**: 関数フィールドとして書きます（VBA からは `obj.Method(args)` で呼ばれる）。
- **デフォルトプロパティ（`obj(i)` 形式）**: オブジェクト本体を **関数** にして、追加のプロパティを生やします。
- **`For Each` 対応**: 関数 / オブジェクトに `[Symbol.iterator]` を実装する。
- **参照設定 (`Dim x As ClassName` / `New ClassName`) 対応**: factory が返すオブジェクトに `__className__: 'クラス名'` を含めると、その名前でも自動的に別名登録されます。

### サンプル: `VBScript.RegExp` のモック

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

### 登録例

```typescript
import { VBARunner, assert } from '../../test-libs/test-runner';
import { createRegExpMock } from '../../test-libs/regexp-mock';

const vbaRunner = new VBARunner('macro.bas');
vbaRunner.registerComObject(createRegExpMock);
```

### 組み込み外部オブジェクト

組み込みで使用可能なライブラリの一覧と差し替え方法は、前セクション「[外部ライブラリの参照設定と `New ProjectName.ClassName`](#外部ライブラリの参照設定と-new-projectnameclassname)」を参照してください。

### Auto-Instantiation 仕様（`Dim x As New ClassName`）

- **遅延インスタンス化**: 宣言時点では実際のインスタンスは作成されず、最初のメンバー参照やメソッド呼び出しで自動生成される
- **`Nothing` 後の自動再生成**: `Set x = Nothing` した後でも、再度参照すると新しいインスタンスが自動生成される
- **`x Is Nothing` は常に False**: auto-instance 変数は参照前でも `Nothing` ではないと判定される

---

## Excel オブジェクトのモック化（Sheets / Range / Application）

VBA マクロの中でも **Excel オブジェクトへの依存**（`Sheets(...)`, `Range(...)`, `Application.WorksheetFunction` など）は頻出です。これらはこのエンジン上に Excel は存在しないため、テスト時にはモックが必要です。

### 推奨アプローチ：依存の分離を優先する

モック化の前に、まず **Excel I/O とドメインロジックを分離**することを強く推奨します（詳細は [`docs/TESTING_STRATEGY.md`](docs/TESTING_STRATEGY.md) / [`docs/REFACTORING_GUIDE.md`](docs/REFACTORING_GUIDE.md)）。

```vb
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

VBA Runner には `src/engine/mock/MockWorksheet.ts` として `MockApplication` と `MockWorksheet` が実装済みです。

```typescript
import { MockApplication } from '../../src/engine/mock/MockWorksheet';
import { Lexer } from '../../src/engine/lexer';
import { Parser } from '../../src/engine/parser';
import { Evaluator } from '../../src/engine/evaluator';

const mockApp = new MockApplication();
mockApp.Sheets('Data').setCellValue('A1:A3', [[10], [20], [30]]);

const ev = new Evaluator(() => {});
ev.evaluateModule(new Parser(new Lexer(vbaCode).tokenize()).parse());

ev.getGlobalEnv().set('Sheets', (name: string) => {
    const ws = mockApp.Sheets(name);
    return { Range: (addr: string) => ws.Range(addr) };
});

const result = ev.callProcedure('GetTotal', []);
```

### MockApplication の拡張

`MockApplication` を継承して必要な機能を追加します。

```typescript
import { MockApplication } from '../../src/engine/mock/MockWorksheet';

class MockApplicationEx extends MockApplication {
    readonly worksheetfunction = {
        vlookup: (lookupVal: any, table: any[][], col: number) => {
            const row = (table as any[][]).find(r => r[0] === lookupVal);
            if (!row) throw Object.assign(new Error('N/A'), { number: 1004 });
            return row[col - 1];
        },
        sum: (...args: any[]) =>
            args.flat(Infinity).reduce((a: number, b: any) => a + Number(b), 0),
    };
}
```

> **注意**: エンジンは VBA の識別子を小文字で解決します（`Application.WorksheetFunction` → `application.worksheetfunction`）。プロパティ名はすべて lowercase で定義してください。

詳細は [`docs/MOCK_GUIDE.md`](docs/MOCK_GUIDE.md) を参照してください。

---

## VBA 実行時の終了ステータスと例外の検知

TypeScript 側から `vbaRunner.run()` または `vbaRunner.eval()` を呼び出した際、VBA の終了状態を以下のように区別して検知できます。

### 1. 正常終了

プロシージャが最後まで実行された、または `Exit Sub/Function` で抜けた場合、戻り値が TypeScript 側に返ります。

### 2. `End` ステートメントによる強制終了

VBA コード内で `End` ステートメントが実行されると、`{ type: 'Terminate' }` というオブジェクトがスローされます。

```typescript
try {
    vbaRunner.eval("End");
} catch (e: any) {
    if (e && e.type === 'Terminate') {
        console.log("End ステートメントによりプログラムが終了しました");
    }
}
```

### 3. 実行時エラー（例外）

未定義の関数呼び出し、ゼロ除算、型不一致などの実行時エラーが発生した場合は、標準的な `Error` オブジェクトがスローされます。

```typescript
try {
    vbaRunner.run("NonExistentSub", []);
} catch (e: any) {
    console.error("実行時エラー:", e.message);
}
```

### 4. `Stop` ステートメント

`Stop` ステートメントは、現在の実装では `console.log` に `STOP Statement encountered` と出力するのみで、プログラムの実行はそのまま継続されます。

### 未初期化変数の暗黙初期化

VBA仕様に従い、宣言されただけで値を設定していない変数は `0` として扱われます（`vbaEmpty`/`null` ではありません）。

```typescript
// VBA: Dim x As Long （代入なし）
assert.strictEqual(vbaRunner.run('GetX', []), 0);
```

---

## ファイル入出力のSandbox方針

VBAコードからのファイル入出力はすべて**Sandboxディレクトリ**内に限定します。

### Sandboxの動作ルール

1. **Sandboxルートの指定**: デフォルトは `sandbox/`（リポジトリルートからの相対パス）です。
2. **Windowsパスの仮想化**: `C:\foo\bar.txt` → `{sandboxRoot}/c/foo/bar.txt` にマッピングされます。
3. **上位ディレクトリへのアクセス禁止**: `../` などを使ってSandboxルートの外へのアクセスを試みた場合はランタイムエラーになります。
4. **相対パスの扱い**: VBAコード内の相対パスはSandboxルートを起点として解決されます。

### パス変換ユーティリティ（テスト用）

```typescript
import { SandboxPath } from '../../test-libs/sandbox';

const sandbox = new SandboxPath('./sandbox');

// 実パス → VBA上の仮想パス（Windows形式）
sandbox.toVirtualPath('./sandbox/c/foo/bar.txt'); // => "C:\\foo\\bar.txt"

// VBA上の仮想パス → 実パス
sandbox.toRealPath('C:\\foo\\bar.txt');             // => "./sandbox/c/foo/bar.txt"
sandbox.toRealPath('D:/data/output.csv');           // => "./sandbox/d/data/output.csv"
```

### 環境変数のSandbox化

VBAの `Environ` 関数による環境変数アクセスも実際のOSの環境変数とは切り離します。

```typescript
// メモリ上で環境変数を定義する例
const sandbox = new SandboxPath('./sandbox', {
  env: {
    USERNAME: 'testuser',
    APPDATA: 'C:\\Users\\testuser\\AppData\\Roaming',
  }
});
```

---

## 仮想ファイルシステム (VFS) の利用

`VBARunner` は `MemoryFileSystem` (VFS) をデフォルトのファイルシステムとして使用します。

- **安全性**: テスト実行によってホスト OS のファイルが誤って削除・上書きされるリスクがゼロ
- **再現性**: テスト実行ごとにクリーンなファイルシステム状態から開始でき、並列実行も容易
- **ブラウザ互換性**: テストコードそのものをブラウザ上でそのまま実行可能

### JavaScript から VFS への直接アクセス (テストデータの準備)

```typescript
import { VBARunner } from './test-libs/test-runner';

const vbaRunner = new VBARunner('source.bas');
const fs = vbaRunner.evaluator.fs;
fs.writeFileSync('/sandbox/c/input.txt', "データ");
vbaRunner.run('ProcessFile', []);
const result = fs.readFileSync('/sandbox/c/output.txt', 'utf-8');
```

VBA の `SandboxPath` が `C:\foo` → `/sandbox/c/foo` と変換するため、パスは `/sandbox/c/` 配下を使うことをお勧めします。

### ファイルシステムの仕様

| 内部処理 | 説明 |
|---------|------|
| ファイルシステム | `MemoryFileSystem` |
| `/sandbox/c/foo` の解釈 | メモリ上のキー |
| `SandboxPath.toRealPath('C:\\foo')` | `/sandbox/c/foo` (仮想) |
| 環境変数隔離 | ✅ 完全隔離 |

---

## 高度な機能と仕様

### 複数モジュールの評価とスコープ

`VBARunner` に対して、ディレクトリを指定して複数のモジュールをまとめてロードできます。

| 拡張子 | 対象モジュール種別 |
|--------|------------------|
| `.bas` | 標準モジュール |
| `.cls` | クラスモジュール |
| `.frm` | フォームモジュール |

> **ファイルロード順**: アルファベット順で確定的にロードされます。モジュールレベルの初期化コード（`Sub`/`Function` 外）がある場合はこの順で実行されます。

### スコープ修飾子（Public / Private / Friend）について

| 呼び出し元 | Private プロシージャ | Public / 修飾子なし |
|-----------|-------------------|-------------------|
| `vbaRunner.run()` / `vbaRunner.eval()`（TypeScript） | ✅ 呼び出し可能 | ✅ 呼び出し可能 |
| **同一モジュール**内の VBA コード | ✅ 呼び出し可能 | ✅ 呼び出し可能 |
| **別モジュール**の VBA コード | ❌ 実行時エラー | ✅ 呼び出し可能 |

- **テストランナーからは `Private` でも直接呼び出せます**。プライベートなヘルパー関数も個別に検証できます。
- `Friend` はパース・保持されますが、実行時のアクセス制御は `Public` と同等です（制御なし）。
- 修飾子なしのプロシージャは VBA 仕様どおり `Public` と同等です。

---

## VBA実行エンジンの動作仕様とインターフェース

### 1. VBAソースのコンパイルと実行 (`evaluate`)

VBAソースコード（`.bas`）をパースして得られたAST（抽象構文木）を順次実行します。

- **宣言の登録**: `Sub`, `Function`, `Property`, `Type`, `Const`, `Dim` などの宣言が出現すると、現在の実行環境（Environment）に登録されます。
- **トップレベル実行**: 本実行エンジンでは**トップレベルでの実行ステートメントを許可**しています（実際のVBA/Excelでは不可）。

### 2. `vbaRunner.run(procedureName, args)` の仕様

- **引数の渡し方**: 第二引数にJavaScriptの配列を渡すと、VBA側のプロシージャ引数にマッピングされます。
- **戻り値**: `Function` または `Property Get` の戻り値をJavaScriptの値として返します。`Sub` の場合は `undefined` を返します。

### 3. `vbaRunner.eval(exprString)` の仕様

1. まず「VBA式（Expression）」としてパース・評価を試みます。成功すればその値を返します。
2. 式としてパースできない場合、自動的に「ステートメント（Statement）」としてパースし、`evaluate` を実行します。

---

## 制限事項

### 1. Boolean 型の扱い

VBAの `Boolean` 型は `VbaBoolean` というラッパークラスで実装しています。`vbaRunner.run()` 等の戻り値を直接 JavaScript の `if` 文で判定すると、オブジェクトであるため **`False` であっても `true` と評価されます**。JS側で真偽判定を行う場合は `result.valueOf() !== 0` とするか、`assert.isTrue(result)` を使用してください。

### 2. 数値型（Integer / Long / Double）の未分化

**実行時の値はすべて JavaScript の `number` (64bit float) として扱います**。

- `TypeName(10)` → `"Integer"`、`TypeName(40000)` → `"Long"`、`TypeName(1.5)` → `"Double"`
- 式の結果: `TypeName(10 + 20)` → `"Double"`（演算結果はデフォルトで Double 扱い）
- Integer / Long のオーバーフロー検出や型キャストによる切り捨ては行われません。

### 3. コレクション・配列のインデックス

- **配列**: VBA インデックスを JavaScript の配列インデックスにそのままマップします。JS 配列の `length` は常に `upper + 1` で、`[0]`〜`[lower-1]` は `undefined`、`[lower]`〜`[upper]` に値が入ります。
- **コレクション**: `Scripting.Dictionary` は JS の `Map` をベースにしたプロキシオブジェクトとして動作します。

### 4. 外部コマンドの実行 (Shell)

`Shell` 関数は外部コマンドを実際に実行しません。実行しようとしたコマンドの内容をログに記録し、固定のタスクID（1）を返します。

---

## VBA でテストを書き、実環境で実行する

VBA で直接テストコードを記述し、実 VBA 環境（Excel など）で実行するための仕組みを提供しています。**Excel マクロ環境で直接動作を検証したい場合**に適しています。

### VBA テストソースの構成ルール

`tests/vba/` ディレクトリに VBA ファイルを配置します。

| 種類 | 命名 | 例 |
|---|---|---|
| テストプロシージャを含む `.bas` | `<名前>Test.bas`（**Test サフィックス**） | `CurrencyOperationsTest.bas` |
| テスト専用サポートクラス `.cls` | 短い名前（`Test` サフィックスなし） | `RefA.cls`, `Helper.cls` |
| 共通インフラクラス `.cls` | 用途がわかる名前 | `AssertHelper.cls` |
| 生成されたランナー `.bas` | `<名前>Test_runner.bas` | `CurrencyOperationsTest_runner.bas` |

> **注意**: ファイル名から拡張子を除いた文字列が VBA のモジュール名になります。VBA のモジュール名は **31 文字以内** でなければなりません。

1 つのテストで複数のクラスファイルが必要になる場合は、サブディレクトリを作成してファイル名を短くします。

### VBA テストコードの書き方

テストプロシージャは `Sub Test_<機能名>(assert)` の形式で定義し、`assert` 引数に `AssertHelper` インスタンスが渡されます。

```vb
' tests/vba/CurrencyOperationsTest.bas
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
```

#### SetUp と TearDown

テストランナーは **`SetUp`** と **`TearDown`** を自動検出し、各テストの前後に呼び出します。

```vb
Sub SetUp()
    testCounter = 0
    testState = "initialized"
End Sub

Sub TearDown()
    testCounter = 0
    testState = ""
End Sub
```

### テストランナーの自動生成

`test-libs/vba-test-generator.ts` ツールを使用して、VBA テストファイル内のすべての `Test_*` プロシージャを検出し、実 VBA 環境で動作するテストランナー Sub を自動生成できます。

```bash
# 単一ファイルのテストランナーを生成
npx tsx test-libs/vba-test-generator.ts tests/vba/CurrencyOperationsTest.bas tests/vba/CurrencyOperationsTest_runner.bas

# ディレクトリ内のすべてのテストファイルを処理
npx tsx test-libs/vba-test-generator.ts --dir tests/vba
```

### 実 VBA 環境でテストを実行

1. Excel を開き、VBA エディターを起動（Alt + F11）
2. 生成されたテストファイルと `*_runner.bas` をインポート
3. `RunAllTests` Sub を実行（F5 キー）
4. メッセージボックスとイミディエイトウィンドウで結果を確認

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
