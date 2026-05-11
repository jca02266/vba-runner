# VBA Execution Environment

Excelがない環境で、作成したVBAコードの動作確認、リファクタリング、および自動テスト（ユニットテスト）を実行するためのプロジェクト（VBAパーサーおよびAST評価器）です。CLIによるTypeScriptのテストランナー環境を備えています。おまけでReactのWeb UIも用意しています。

## 本プロジェクトの目的
- **環境非依存での動作確認**: Excel (Windows/Mac) に依存せず、モダンなブラウザ上またはサーバー上(Node.js等)で直接VBAの構文とロジックを実行・検証できます。
- **リファクタリングの支援**: 巨大なVBAコードベースから純粋な関数・サブルーチンを安全に切り出し、モジュールを整理するための検証基盤を提供します。
- **ユニットテストの実行**: TypeScriptを利用したテストランナーを通じて、抽出されたVBA関数に対して直接モックデータやアサーションを評価し、プログラムによるテスト自動化を可能にします。

上記より、VS Codeなどのモダンな環境で（AIによる支援を受けながら）VBAソースのリファクタリングとテストの実行を可能にします。

## 基本方針：環境依存操作のダミー化

本プロジェクトは**リファクタリングとユニットテストの実行**を主目的としています。そのため、実行環境（OSやファイルシステム）に副作用を及ぼす、または環境に依存する以下の操作は、原則として**サンドボックス化またはダミー（モック）実装**で行います。

- **外部コマンドの実行 (`Shell`)**: 実際のコマンドは実行せず、ログ出力と成功ステータスの返却のみを行います。
- **ファイル操作 (`Open`, `Kill` 等)**: 指定されたサンドボックスディレクトリ (`workspace/`) 内に制限されます。
- **環境変数 (`Environ`)**: OSの実環境変数にはアクセスせず、仮想的な環境変数（`.env` 等）を使用します。
- **レジストリ (`GetSetting`, `SaveSetting`)**: Windowsレジストリにはアクセスせず、メモリ上のマップ等による仮想レジストリを使用します。
- **外部オブジェクト・ネットワーク (`CreateObject`, `MSXML2.XMLHTTP`)**: 主要なライブラリをモックとして提供し、実際のCOMオブジェクト生成やネットワーク通信は行いません。
- **ユーザー対話 (`MsgBox`, `InputBox`)**: 実行をブロックせず、ログ出力とデフォルト値の返却のみを行います。

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
 │   └── tests/ts/               
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
- `tests/ts/` - 汎用のテストランナーユーティリティ。

また、以下にVBA評価器自体のテストが配置されています。

- `tests/spec/` - VBAの仕様を確認するための各機能のテスト。

## テストの書き方

VBAソースのテストを記述する際は、テストランナー `vba-compiler/tests/ts/test-runner.ts` を使用します。

> [!TIP]
> 実際にコードを書く前に、ブラウザ上で動作を確認したい場合は [Web UI デモサイト](https://vba-web-compiler.netlify.app/) をご利用ください。`Debug.Print` の結果や構文チェックを即座に試すことができます。

### 1. `eval`: VBAの式やコード断片をその場で評価する

VBAの構文をそのまま文字列として渡し、評価結果を取得します。最も手軽にVBAエンジンを試すことができる方法です。

```typescript
import { VBATest } from '../../tests/ts/test-runner';
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
import { assert } from '../../tests/ts/test-runner';

assert.isTrue(vbaTest.eval('1 = 1'),        '等値比較');
assert.isFalse(vbaTest.eval('IsNull(123)'), 'IsNull(123)');

// assert.strictEqual で書く場合は vbaTrue/vbaFalse をインポートする必要がある
// import { vbaTrue, vbaFalse } from '../../src/compiler/evaluator';
// assert.strictEqual(result, vbaTrue);
```

### 配列

VBAでの宣言が `Dim arr(N)` の場合、JavaScript側では要素数が `N + 1` （インデックス 0〜N）の配列として返ります。

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

1. **Sandboxルートの指定**: Evaluatorの初期化時にSandboxルートディレクトリを指定します。デフォルトは `workspace/`（プロジェクトルートからの相対パス）です。

2. **Windowsパスの仮想化**: VBAコード内の絶対パスはドライブレターをサブディレクトリとしてSandboxルート配下にマッピングします。
   - `C:\foo\bar.txt` → `{sandboxRoot}/c/foo/bar.txt`
   - `D:\data\output.csv` → `{sandboxRoot}/d/data/output.csv`
   - ドライブレター（`C:`、`D:` など）は小文字に正規化してSandboxルート配下のディレクトリとして扱います。

3. **上位ディレクトリへのアクセス禁止**: `../` などを使ってSandboxルートの外へのアクセスを試みた場合はランタイムエラーになります。

4. **相対パスの扱い**: VBAコード内の相対パスはSandboxルートを起点として解決されます。

### パス変換ユーティリティ（テスト用）

`tests/ts/sandbox.ts` に実パスと仮想パス（VBAから見えるWindowsパス）を相互変換するユーティリティを提供します。

```typescript
import { SandboxPath } from '../../tests/ts/sandbox';

const sandbox = new SandboxPath('./workspace');

// 実パス → VBA上の仮想パス（Windows形式）
sandbox.toVirtualPath('./workspace/c/foo/bar.txt'); // => "C:\\foo\\bar.txt"

// VBA上の仮想パス → 実パス
sandbox.toRealPath('C:\\foo\\bar.txt');             // => "./workspace/c/foo/bar.txt"
sandbox.toRealPath('C:/foo/bar.txt');               // => "./workspace/c/foo/bar.txt"
sandbox.toRealPath('D:/data/output.csv');           // => "./workspace/d/data/output.csv"
```

テストコードでは、VBAコードに渡すパスをこのユーティリティで組み立てることで、実ファイルシステムへの影響を限定しながらファイル操作の動作を検証できます。

### 環境変数のSandbox化

VBAの `Environ` 関数による環境変数アクセスも実際のOSの環境変数とは切り離します。

- **定義方法**: Sandboxの設定ファイル（`{sandboxRoot}/.env`）に `KEY=VALUE` 形式で記述するか、Evaluatorの初期化時にメモリ上のマップとして渡します。
- **実環境変数へのアクセス禁止**: VBAコードから `Environ("PATH")` などを呼び出しても、OSの実際の環境変数は参照されません。Sandbox内に定義のないキーは空文字列 `""` を返します。
- **優先順位**: メモリ上の定義 → `{sandboxRoot}/.env` の順に参照します。

```typescript
// メモリ上で環境変数を定義する例
const sandbox = new SandboxPath('./workspace', {
  env: {
    USERNAME: 'testuser',
    APPDATA: 'C:\\Users\\testuser\\AppData\\Roaming',
  }
});
```

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

CLIからTypeScriptのテストランナーを利用し、`sample/src/vba` 以下のVBAモジュール内の関数を直接評価・検証します。

```bash
# バンドルしてテスト（AST構築・検証）を実行
npx esbuild sample/tests/ts/TaskScheduler_Core.test.ts --bundle --outfile=sample/tests/ts/TaskScheduler_Core.test.cjs --platform=node && node sample/tests/ts/TaskScheduler_Core.test.cjs
```

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
VBAには `Integer` (16bit), `Long` (32bit), `Double` (64bit float) 等の厳密な数値型がありますが、本エンジンでは実行速度とシンプルさを優先し、**すべての数値を JavaScript の `number` (64bit float) として扱います**。

- **影響**: `TypeName(10)` は本来 `"Integer"` を返すべきですが、本エンジンでは一律に `"Double"` を返します。
- **今後の対応**: 現在、`Integer` や `Long` の厳密な区別の優先度は低く設定されています。計算精度が必要な場合は、JSの数値精度（Double相当）の範囲内で動作することに留意してください。

### 3. コレクション・配列のインデックス
- **配列**: VBAでは `Option Base 1` や `Dim arr(1 To 5)` が可能ですが、本エンジンでは原則として 0-indexed な JS 配列として扱います。
- **コレクション**: `Scripting.Dictionary` は JS の `Map` をベースにしたプロキシオブジェクトとして動作します。

### 4. 外部コマンドの実行 (Shell)
- **制限**: セキュリティおよびユニットテストへの集中を目的とし、`Shell` 関数は外部コマンドを実際に実行しません。
- **挙動**: 実行しようとしたコマンドの内容を `Debug.Print` 相当の出力としてログに記録し、固定のタスクID（1）を返します。
