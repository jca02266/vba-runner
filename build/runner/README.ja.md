# vba-runner

Excel 不要で VBA コードを実行・テストできる TypeScript 実装の VBA 実行エンジンです。
テストランナーとして使う場合は TypeScript からVBAファイルをロードして関数を呼び出せます。
また、静的解析・整形・構文チェックの CLI ツールも同梱しています。

> 対象: パッケージ利用者
>
> 目的別のガイド一覧: [ドキュメントガイド](../../docs/README.md)

## インストール

```bash
npm install vba-runner
```

## テストの書き方

> [!TIP]
> 実際にコードを書く前に、ブラウザ上で動作を確認したい場合は [Web UI デモサイト](https://vba-web-runner.netlify.app/) をご利用ください。`Debug.Print` の結果や構文チェックを即座に試すことができます。

### 1. `eval`: VBAの式やコード断片をその場で評価する

VBAの構文をそのまま文字列として渡し、評価結果を取得します。最も手軽にVBAエンジンを試すことができる方法です。

```typescript
import { VBARunner } from 'vba-runner';
const vbaRunner = new VBARunner(); // 空の環境を作成

// VBAの計算式を直接評価
const sum = vbaRunner.eval("1 + 2 + 3"); // => 6

// 変数宣言や代入を含む、複数行の処理も実行可能
vbaRunner.eval("Dim x : x = 10 : Debug.Print x * 2");
```

### 2. `run`: ロードしたソース内の関数を実行する

既存のVBAファイル（`.bas`）をロードし、そこに定義されたプロシージャを引数を指定して呼び出します。複雑なビジネスロジックのユニットテストに最適です。

```typescript
import { VBARunner, assert } from 'vba-runner';

// 1. テスト対象のVBAファイルをロード
const vbaRunner = new VBARunner('src/vba/Sample.bas');

// 2. 関数名を指定して実行（第2引数はJavaScriptの配列として引数を渡す）
const result1 = vbaRunner.run('Add', [1, 2]);
const result2 = vbaRunner.run('Multiply', [result1, 2]);

// 3. 結果をアサート
assert.strictEqual(result1, 3);
assert.strictEqual(result2, 6);
```

### 3. glob パターンでファイルを選択

glob パターン文字列（またはパターンの配列）を渡すと、マッチするファイルをロードします。
bash と同様に `*`、`**`（再帰）、`{a,b}`（ブレース展開）が使えます。

```typescript
import { VBARunner, assert } from 'vba-runner';

// 単一 glob パターン
const vbaRunner = new VBARunner('src/vba/*.{bas,cls}');

// サブディレクトリを再帰的に検索
const vbaRunner2 = new VBARunner('src/**/*.bas');

// パターンとリテラルパスの混在配列
const vbaRunner3 = new VBARunner([
    'src/vba/core/*.bas',
    'src/vba/models/Account.cls',
]);
```

パターンは `process.cwd()` 基準で解決されます。mock ディレクトリスキャンは行われません（mock はディレクトリ読み込み時の機能）。

### 4. ファイルを配列で個別指定

ロードするファイルを個別に選択したい場合は、リテラルパスの配列を渡します。

```typescript
import { VBARunner, assert } from 'vba-runner';

const vbaRunner = new VBARunner([
    'src/vba/Calc.cls',
    'src/vba/Utils.bas',
]);

const result = vbaRunner.run('DoubleIt', [21]);
assert.strictEqual(result, 42);
```

### 5. 複数ソースの一括ロード

VBAファイル（`.bas`）およびクラスモジュール（`.cls`）を格納したディレクトリを指定すれば、
配下のソースファイルをすべて読み込みます。規模の大きなVBAプロジェクトではこの使い方になります。

```typescript
import { VBARunner, assert } from 'vba-runner';

// 1. テスト対象のVBAディレクトリをロード（.bas と .cls を自動検出）
const vbaRunner = new VBARunner('src/vba');

// 2. 関数名を指定して実行
const result = vbaRunner.run('CalcTotal', [100, 200, 300]);
assert.strictEqual(result, 600);
```

クラスモジュール（`.cls`）はそのまま使えます。VBA標準の `.cls` ヘッダーは自動的に除去されます。
`Class_Initialize` / `Class_Terminate`、`Property Get/Let/Set`、`Private`/`Public` メンバーも動作します。

```typescript
// src/vba/BankAccount.cls
// VERSION 1.0 CLASS
// ...
// Attribute VB_Name = "BankAccount"
// Private m_balance As Double
// ...

const vbaRunner = new VBARunner('src/vba');
vbaRunner.eval('Dim acct As New BankAccount');
vbaRunner.eval('acct.Deposit 1000');
assert.strictEqual(vbaRunner.eval('acct.Balance'), 1000);
```

> **`eval()` と `run()` の呼び出し間で状態は持続します。** 同じ `VBARunner` インスタンス内では、
> `eval()` で宣言した変数やオブジェクトが後続の `eval()` / `run()` でも参照できます。
> テストのセットアップをステップごとに分けて記述するのに便利です。

### 4. Boolean値を比較する

`run()` と `eval()` は VBA の `Boolean` 戻り値を自動的に JS の `true`/`false` に変換します。
標準の assert がそのまま使えます。

```typescript
import { VBARunner, assert } from 'vba-runner';

const vbaRunner = new VBARunner('src/vba/Sample.bas');
assert.strictEqual(vbaRunner.run('IsPositive', [5]), true);  // OK
```

### 5. Excel依存オブジェクト(Application / ActiveSheet など)をモックする

`ActiveSheet.Range(...)` や `Cells(...)` を使うVBAコードは、`excelStub` オプションを
有効にすると組み込みのモックでテストできます。

```typescript
import { VBARunner, assert } from 'vba-runner';

const vbaRunner = new VBARunner('src/vba/Sheet1Logic.bas', { excelStub: true });

// テスト前にセルの初期値を設定
vbaRunner.excelStub.ActiveSheet.setCellValue('A1', 100);

vbaRunner.run('DoubleA1ToB1', []);

// セルの値を検証
assert.strictEqual(vbaRunner.excelStub.ActiveSheet.getCellValue('B1'), 200);
```

> **注意1: `Cells(row, col)` も `Range(アドレス)` と同じストレージを共有する。**
> VBAコード側で`Cells(13, 1)`のように数値アドレッシングを使った場合も、内部的には`Range("A13")`相当に
> 変換されて`setCellValue`/`getCellValue`と同じセルを指す。テストの初期値設定は`setCellValue('A1', ...)`
> のような文字列アドレスで行えば、`Cells`・`Range`どちらでアクセスするコードでも一致する。

組み込みモックは `Value` の読み書きのみサポートしています。`Interior.Color` などの
書式設定は値を保持しない no-op、`Application.OnKey` / `Application.OnTime` は未実装です
(呼び出すとエラーになります)。これらに依存するコードをテストしたい場合は、
**`Application` を丸ごと差し替えるのではなく、モックを拡張してください**。
丸ごと差し替えると、対象コードが同時に使う`ActiveSheet`/`Sheets`/`Range`が失われます。

> **注意2: `Application.OnTime`に1秒未満のオフセットを渡すコード(`Now + 0.4/86400`等)は実機で
> 無限ループに陥る場合がある。** `Now()`は秒未満を切り捨てるため、計算した目標時刻が既に過去になる
> ことがあり、`OnTime`は過去の時刻だと即座に発火する。タイマー処理が自己再スケジュールする実装だと、
> これが延々と繰り返されてUIスレッドを埋め尽くす(CPU使用率は低いまま「応答なし」になるため気づきにくい)。
> オフセットは必ず1秒以上にすること。

最もきれいな方法は、`MockApplication`(このパッケージからexportされている)をサブクラス化し、
`excelStub`に`true`の代わりにインスタンスを渡すことです:

```typescript
import { VBARunner, MockApplication, assert } from 'vba-runner';

class AppWithOnKey extends MockApplication {
  onKeyLog: string[] = [];
  OnKey(key: string, procedureName?: string) {
    this.onKeyLog.push(`${key}=${procedureName ?? ''}`);
  }
}

const vbaRunner = new VBARunner('src/vba/KeyHandler.bas', { excelStub: new AppWithOnKey() });

vbaRunner.run('SetupKeyHandlers', []);
assert.ok(vbaRunner.excelStub.onKeyLog.length > 0);
```

再利用可能なクラスが不要なら、構築後に`excelStub`インスタンスへ直接メソッドを
追加する方法もあります。`vbaRunner.excelStub`は`Application`/`ActiveSheet`として
登録済みのオブジェクトそのものなので、これを拡張する形でも動きます:

```typescript
const vbaRunner = new VBARunner('src/vba/KeyHandler.bas', { excelStub: true });

let registered: [string, string] | null = null;
(vbaRunner.excelStub as any).OnKey = (key: string, procedureName: string) => {
  registered = [key, procedureName];
};

vbaRunner.run('SetupKeyHandlers', []);
assert.ok(registered !== null);
```

`vbaRunner.evaluator.setBuiltinOverride(name, value)` は、`Application` などの
組み込みオブジェクト自体を別物に丸ごと差し替えたい場合にのみ使ってください。
この方法は(対象コードがその名前経由で `ActiveSheet` 等に一切アクセスしない場合を除き)
デフォルトの Excel スタブが提供する残りの機能を失わせる点に注意してください。

任意の名前の変数・定数を直接注入したい場合は `vbaRunner.set(name, value)` も使えます。

詳細なモック実装パターンは [`docs/MOCK_GUIDE.md`](https://github.com/jca02266/vba-runner/blob/main/docs/MOCK_GUIDE.md) を参照してください。

### 6. 静かなモード / `Debug.Print` 出力の分離

デフォルトでは `run()` は呼び出しごとに `[PASS] ProcName(...) -> result (Nms)` を
ログ出力し、`Debug.Print` も `console.log` に出力されるため、両者が同じ標準出力に
混在します。`quiet: true` を渡すと `[PASS]` ログを抑制でき、`onPrint` を渡すと
`Debug.Print` の出力先を独立して切り替えられます（例: stderr や配列への収集）。

```typescript
import { VBARunner } from 'vba-runner';

const vbaRunner = new VBARunner('src/vba/Sample.bas', {
  quiet: true,                                    // [PASS] ログを抑制
  onPrint: (s) => process.stderr.write(s + '\n'),  // Debug.Print → stderr
});

vbaRunner.run('SeedGlider', [1, 1]); // [PASS] ログは出力されない
```

### 8. 仮想ファイルシステム（`vbaRunner.fs`）

VBA のファイル入出力（`Open`、`Print #`、`Line Input #`、`Close`、`Scripting.FileSystemObject` など）は
**インメモリの仮想ファイルシステム（VFS）**上で動作します。実際のディスクには何も書き込まれません。
`VBARunner` インスタンスごとに独立した VFS が作成されます。

VBA コード内で使用する Windows 形式のパスは自動的に VFS 上のパスに変換されます。

| VBA パス | VFS パス（デフォルト `sandboxRoot: '/sandbox'`） |
|---|---|
| `C:\data\input.txt` | `/sandbox/c/data/input.txt` |
| `D:\report.csv` | `/sandbox/d/report.csv` |
| `relative\path.txt` | `/sandbox/relative/path.txt` |

`vbaRunner.fs` を使って、TypeScript 側からテスト用の入力ファイルを事前配置したり、VBA が書き出したファイルを読み取ったりできます。

```typescript
import { VBARunner } from 'vba-runner';

const vbaRunner = new VBARunner('src/vba/Report.bas');

// VBA コードが読み込む入力ファイルを事前配置
vbaRunner.fs.mkdirSync('/sandbox/c/data', { recursive: true });
vbaRunner.fs.writeFileSync('/sandbox/c/data/input.csv', 'Alice,100\nBob,200\n');

// ファイルを読み込んでサマリーを書き出す VBA プロシージャを実行
vbaRunner.run('GenerateSummary', []);

// VBA が書き出したファイルを TypeScript 側で読み取る
const output = vbaRunner.fs.readFileSync('/sandbox/c/data/summary.txt', 'utf8');
console.log(output);
```

VFS のルートパスを変更したい場合は、コンストラクタに `sandboxRoot` を指定します。

```typescript
const vbaRunner = new VBARunner('src/vba/Report.bas', { sandboxRoot: '/myapp' });
// C:\file.txt が /myapp/c/file.txt にマッピングされる
vbaRunner.fs.writeFileSync('/myapp/c/file.txt', 'データ');
```

> `Do While Not ts.AtEndOfStream`（FSO TextStream）、`Do While Not EOF(n)`（ネイティブ `Open`）の
> いずれのパターンも VFS 上で正しく動作します。

### 9. `ByRef` の出力パラメーター

VBA の既定の引数渡しは `ByRef` です。呼び出した `Sub`/`Function` が `ByRef`
パラメーターに代入した場合、`run()` は呼び出し元が渡した JS 配列にその最終値を
書き戻します。「ステータス + メッセージ」パターンによく使われます。

```typescript
import { VBARunner } from 'vba-runner';

const vbaRunner = new VBARunner('src/vba/Sample.bas');

const args: any[] = [/* plate */ 'ABC-123', /* outMessage */ ''];
vbaRunner.run('ParkCar', args);
console.log(args[1]); // 呼び出し後、Sub の ByRef outMessage パラメーターの値が入る
```

`ByVal` パラメーターは書き戻されません（呼び出し元の元の値が維持されます）。

### 10. 条件付きコンパイル（`#If` / `#Const`）

vba-runner は VBA の条件付きコンパイルディレクティブ（`#If`、`#Else`、`#ElseIf`、`#End If`、`#Const`）をサポートしています。

**デフォルトのコンパイラー定数**は、現代的な 64bit Windows VBA 環境（Office 2010+）を模擬します。

| 定数 | デフォルト | 意味 |
|---|---|---|
| `VBA7` | `True` (-1) | VBA バージョン 7（Office 2010+） |
| `Win64` | `True` (-1) | 64bit プロセス |
| `Win32` | `True` (-1) | Windows プラットフォーム（64bit 環境でも常に True） |
| `Mac` | `False` (0) | Mac プラットフォーム |

これらのデフォルトにより、`#If VBA7 Then` や `#If Win64 Then` で囲まれたコードがデフォルトで**有効**になり、最新の Office 環境に合わせた動作になります。

**コンパイラー定数の上書き** — コンストラクターに `compilerConstants` を渡すことで、別の環境をシミュレートできます。

```typescript
// 32bit Office（Office 2010 以前、または 64bit Windows 上の 32bit Office）をシミュレート
const vbaRunner = new VBARunner('src/vba/Module.bas', {
  compilerConstants: { VBA7: 0, Win64: 0 },
});
```

```typescript
// Mac VBA をシミュレート
const vbaRunner = new VBARunner('src/vba/Module.bas', {
  compilerConstants: { Mac: -1, Win32: 0, Win64: 0 },
});
```

**注意:** VBA ソースファイル内の `#Const` ディレクティブは、コンストラクターに渡した `compilerConstants` より優先されます（VBA 仕様準拠）。

## CLI ツール

`vba-runner` パッケージは以下の CLI ツールを提供します。

| コマンド | 説明 |
|---|---|
| `vba-run <file.bas>` | VBAファイルを実行し、Debug.Print の出力を表示する |
| `vba-analyzer <file.bas>` | VBAコードの静的解析（アウトライン・参照数・重複検出など） |
| `vba-formatter <file.bas>` | VBAコードの整形（インデント・スペースの統一） |
| `vba-parse-check <file.bas>` | VBAコードの構文チェック（パースエラーの検出） |

```bash
# VBAファイルの静的解析（アウトライン表示）
vba-analyzer --outline src/vba/Module1.bas

# 重複ブロックの検出
vba-analyzer --diff src/vba/

# 構文チェック
vba-parse-check src/vba/Module1.bas
```

## 詳細ドキュメント

- [ドキュメントガイド](../../docs/README.md) — 目的別の共通ハブ（利用者／開発者）
- [REFERENCE.md](https://github.com/jca02266/vba-runner/blob/main/REFERENCE.md) — 型システム・モック登録・Sandbox方針・VFS など詳細仕様
- [README.md](https://github.com/jca02266/vba-runner/blob/main/README.md) — プロジェクト概要
- [FOR_AI.md](https://github.com/jca02266/vba-runner/blob/main/FOR_AI.md) — AIによるリファクタリング支援ガイド
