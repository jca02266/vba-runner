# VBA Execution Environment

Excelがない環境で、作成したVBAコードの動作確認、リファクタリング、および自動テスト（ユニットテスト）を実行するためのプロジェクト（VBAパーサーおよびAST評価器）です。CLIによるTypeScriptのテストランナー環境を備えています。おまけでReactのWeb UIも用意しています。

## 本プロジェクトの目的
- **環境非依存での動作確認**: Excel (Windows/Mac) に依存せず、モダンなブラウザ上で直接VBAの構文とロジックを実行・検証できます。
- **リファクタリングの支援**: 巨大なVBAコードベースから純粋な関数・サブルーチンを安全に切り出し、モジュールを整理するための検証基盤を提供します。
- **ユニットテストの実行**: TypeScriptを利用したテストランナーを通じて、抽出されたVBA関数に対して直接モックデータやアサーションを評価し、プログラムによるテスト自動化を可能にします。

上記より、VS Codeなどのモダンな環境で(AIによる支援を受けながら)VBAソースのリファクタリングとテストの実行を可能にします。

## ディレクトリ構成とVBAサンプルコード
本プロジェクトでは、巨大なVBAマクロを「テスト可能な単位」にリファクタリングする事例として、以下の構成を採用しています。

- `src/compiler/` - TypeScriptで書かれたVBA用Lexer、Parser、Evaluatorのコアエンジン群。
- `sample/src/vba/` - 本番稼働を想定したプロダクションコード。
  - `TaskScheduler.vba`: 処理の見通しが良くなったリファクタリング**後**のメインルーチン。
  - `TaskScheduler_Core.vba`: `TaskScheduler.vba` から抽出された、純粋なビジネスロジック関数モジュール（**ユニットテストの対象**）。
- `sample/src/vba_legacy/` - 参照用のレガシーコード。
  - `TaskScheduler_v1.vba`: リファクタリング**前**の、機能が密結合した巨大なオリジナルソース。
- `sample/tests/ts/` - Node.js 上でVBAスクリプトをAST評価し、自動テストを走らせるTypeScriptスクリプト群。
  - `TaskScheduler_Core.test.ts`: `TaskScheduler_Core.vba` 内の各関数の振る舞いを検証するユニットテスト。
- `tests/ts/` - 汎用のテストランナーユーティリティ。

## VBATest — テストランナーの使い方

`tests/ts/test-runner.ts` の `VBATest` クラスは、単一ファイルとディレクトリ（複数ファイル）の両方をロードできます。

### 単一ファイルを指定する（従来の使い方）

```typescript
import { VBATest } from '../../tests/ts/test-runner';

const vbaTest = new VBATest('sample/src/vba/TaskScheduler_Core.vba');
vbaTest.run('CalcBaseStartIdx', [1, 10, 1.0]);
```

### ディレクトリを指定して複数ファイルをまとめてロードする

ディレクトリを渡すと、配下の `.vba` / `.cls` / `.frm` ファイルをアルファベット順にすべてロードし、
**同一スコープ**に登録します。どのファイルに定義されたプロシージャでも相互に呼び出せます。

```typescript
const vbaTest = new VBATest('src/vba/');   // ディレクトリを指定

// MathUtils.vba のプロシージャ
vbaTest.run('Add', [3, 4]);          // => 7

// StringUtils.cls のプロシージャ
vbaTest.run('Greet', ['World']);     // => "Hello, World!"

// Main.frm のプロシージャが他ファイルの Add / Greet を呼び出す
vbaTest.run('CalcAndGreet', ['Alice', 10, 20]);
```

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

## テスト環境での `run` と `eval` の使い分け
TypeScript側からVBAのロジックを呼び出す際、用途に応じて以下の2つのメソッドを利用します。

- **`vbaTest.run(procedureName, args)`**:
  - **用途**: VBAのサブルーチン（`Sub`）や関数（`Function`）に、TypeScript側からの配列引数を渡して実行します。
  - **特徴**: トップレベルから直接プロシージャ名を指定して呼び出すための「標準的」な実行メソッドです。
  - **例**: `const result = vbaTest.run("AddNumbers", [100, 200]);`

- **`vbaTest.eval(expressionString)`**:
  - **用途**: VBAの「式（Expression）」や「文（Statement）」を文字列として渡し、その場でVBAとして評価（パース）して実行します。
  - **特徴**: 
    - `AddNumbers(3, 4)` のように、VBAとして有効な式を評価し、その戻り値をダイレクトに取得する場合に向いています。
    - また、`MainLoop` のように括弧なしで「文」として実行されるVBA特有の呼び出しも自動で判別して実行します（この場合の戻り値は `undefined` となります）。
  - **例**: `const result = vbaTest.eval("AddNumbers(3, 4)");` または `vbaTest.eval("MainLoop");`

## テストコードで使う固有型と値

VBAの型システムはJavaScriptと異なるため、`run` / `eval` の戻り値をアサートする際に以下の対応表を参照してください。

### スカラー値の対応

| VBA の値          | JavaScriptの型・値    | 説明                                                       |
|-------------------|-----------------------|------------------------------------------------------------|
| `Empty`           | `null`（`EmptyVBA`）  | 未初期化のスカラー変数。`evaluator.ts`の`EmptyVBA`定数と同値 |
| `Nothing`         | `null`                | 未設定のオブジェクト参照                                   |
| `True` / `False`  | `true` / `false`      | VBAブール値はJSブール値に直接対応                          |
| 数値              | `number`              | 整数・浮動小数点とも `number` で返る                       |
| 文字列            | `string`              | ダブルクォート除去済みの生文字列                           |
| `Null`            | ※未実装               | SQL由来の3値論理Null（現バージョンでは扱わない）            |

```typescript
import { EmptyVBA } from '../../src/compiler/evaluator';

assert.strictEqual(result, EmptyVBA);   // VBAのEmpty（null）と比較
assert.strictEqual(result, null);       // 同上（直接nullでも可）
```

### 配列（`Dim arr(N)`）

VBAの`Dim arr(N)`はインデックス`0`〜`N`の配列（要素数N+1）として返ります。
JavaScript配列として受け取り、**インデックスは1始まり**で使うのがVBAの慣例です（インデックス0は未使用）。

```typescript
const arr = vbaTest.run('GetArray', []) as any[];
assert.strictEqual(arr[1], 'first');  // VBAのarr(1)はJS配列のインデックス1
assert.strictEqual(arr.length, 6);    // Dim arr(5) → 要素数6（0〜5）
```

### Scripting.Dictionary

`CreateObject("Scripting.Dictionary")` が返すオブジェクトの構造:

```typescript
interface VbaDict {
    __isVbaDict__: true;
    __map__: Map<string, any>;  // 内部のJS Map（テストでの直接参照用）
    add(key: string, value: any): void;
    exists(key: string): boolean;
    items(): any[];
    keys(): string[];
}
```

テストコードでは `__map__` を通じて内容を検証できます:

```typescript
const dict = vbaTest.run('BuildDict', []) as any;
assert.strictEqual(dict.__map__.get('key1'), 'value1');
assert.strictEqual(dict.exists('key1'), true);
```

### Sub と Function の戻り値

| プロシージャ種別 | `run` の戻り値 | `eval` の戻り値 |
|-----------------|---------------|----------------|
| `Sub`（引数あり）| `undefined`   | ——（`run`を使う）|
| `Function`（引数あり）| 戻り値     | ——（`run`を使う）|
| `Function`（引数なし）| 戻り値     | 戻り値         |
| VBA式（`eval`用）| ——           | 評価結果        |

### 未初期化変数の暗黙初期化

VBA仕様に従い、宣言されただけで値を設定していない変数は `0` として扱われます（`EmptyVBA`/`null` ではありません）。テストで「初期化されていないこと」を確認したい場合は実際の返り値（`0` または `""`）を確認してください。

```typescript
// VBA: Dim x As Long （代入なし）
assert.strictEqual(vbaTest.run('GetX', []), 0);
```

### `EmptyVBA`のインポート

```typescript
import { EmptyVBA } from '../../src/compiler/evaluator';
// または
import { EmptyVBA } from '../../../src/compiler/evaluator'; // tests/spec/ からの相対パス
```

## 自動テストの実行
CLIからTypeScriptのテストランナーを利用し、`TaskScheduler_Core.vba` 内の抽出された関数を直接評価・検証します。

```bash
# バンドルしてテスト（AST構築・検証）を実行
npx esbuild sample/tests/ts/TaskScheduler_Core.test.ts --bundle --outfile=sample/tests/ts/TaskScheduler_Core.test.cjs --platform=node && node sample/tests/ts/TaskScheduler_Core.test.cjs
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