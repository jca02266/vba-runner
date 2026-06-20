# vba-runner

Excel 不要で VBA コードを実行・テストできる TypeScript 実装の VBA 実行エンジンです。
テストランナーとして使う場合は TypeScript からVBAファイルをロードして関数を呼び出せます。
また、静的解析・整形・構文チェックの CLI ツールも同梱しています。

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

### 3. 複数ソースの一括ロード

VBAファイル（`.bas`）を格納したディレクトリを指定すれば、配下のソースファイルをすべて読み込みます。
規模の大きなVBAプロジェクトではこの使い方になります。

```typescript
import { VBARunner, assert } from 'vba-runner';

// 1. テスト対象のVBAディレクトリをロード
const vbaRunner = new VBARunner('src/vba');

// 2. 関数名を指定して実行
const result = vbaRunner.run('CalcTotal', [100, 200, 300]);
assert.strictEqual(result, 600);
```

### 4. Boolean値を比較する

VBAの `Boolean` を返す関数の結果は、`assert.strictEqual(result, true)` のように素のJSの
`true`/`false` と直接比較すると失敗します(VBAの`Boolean`値はラッパーオブジェクトとして
返るため)。専用の `assert.isTrue` / `assert.isFalse` を使ってください。

```typescript
import { VBARunner, assert } from 'vba-runner';

const vbaRunner = new VBARunner('src/vba/Sample.bas');
const result = vbaRunner.run('IsPositive', [5]);

assert.isTrue(result);                  // OK
// assert.strictEqual(result, true);    // NG: 失敗する
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

組み込みモックは `Value` の読み書きのみサポートしています。`Interior.Color` などの
書式設定は値を保持しない no-op、`Application.OnKey` / `Application.OnTime` は未実装です
(呼び出すとエラーになります)。これらに依存するコードをテストしたい場合は、
`vbaRunner.evaluator.setBuiltinOverride(name, value)` で `Application` などの
組み込みオブジェクト自体を独自のモックに丸ごと差し替えてください。

```typescript
const vbaRunner = new VBARunner('src/vba/KeyHandler.bas');

let registered: [string, string] | null = null;
const customApplication = {
  OnKey: (key: string, procedureName: string) => { registered = [key, procedureName]; },
};
vbaRunner.evaluator.setBuiltinOverride('Application', customApplication);

vbaRunner.run('SetupKeyHandlers', []);
assert.ok(registered !== null);
```

任意の名前の変数・定数を直接注入したい場合は `vbaRunner.set(name, value)` も使えます。

詳細なモック実装パターンは [`docs/MOCK_GUIDE.md`](https://github.com/jca02266/vba-runner/blob/main/docs/MOCK_GUIDE.md) を参照してください。

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

- [REFERENCE.md](../../REFERENCE.md) — 型システム・モック登録・Sandbox方針・VFS など詳細仕様
- [README.md](../../README.md) — プロジェクト概要・クイックスタート
- [FOR_AI.md](../../FOR_AI.md) — AIによるリファクタリング支援ガイド
