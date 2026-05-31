# VBA サンプル - Excel データ処理

## ファイル

| ファイル | 内容 |
|---|---|
| `DataProcessor.bas` | メインの VBA モジュール |

## 機能

1. **サンプルデータ生成** - `Data` シートに商品売上データを自動生成
2. **重複行の除去** - ID が重複している行を取り除いてコピー
3. **カテゴリ別集計** - カテゴリごとに件数・売上合計・平均売上を計算
4. **結果の出力** - 処理結果を `Result` シートに整形して出力

## 使い方

### Excel へのインポート

1. Excel を開き、`Alt + F11` で VBA エディターを起動
2. メニュー「ファイル」→「ファイルのインポート」で `DataProcessor.bas` を選択
3. エディターを閉じる

### 実行

- `Alt + F8` でマクロ一覧を開き `ProcessData` を選択して「実行」
- または VBA エディター上でカーソルを `ProcessData` 内に置いて `F5`

### 実行結果

```
Data シート   : サンプルデータ(9行・重複1件含む)
Result シート : 重複除去済みデータ + カテゴリ別集計
```

## コード構成

```
ProcessData()          ← エントリポイント
├─ CreateSampleData()  ← Data シートにサンプルデータ投入
├─ CopyFilteredData()  ← 重複除去してコピー
├─ SummarizeByCategory() ← カテゴリ別集計
└─ FormatResultSheet() ← 書式整形
```

## 学習ポイント

- `Scripting.Dictionary` を使った重複チェックと集計
- シートの取得・作成(`GetOrCreateSheet`)
- `End(xlUp).Row` による最終行取得
- セルの書式設定(`Interior.Color`, `NumberFormat`)
- `Option Explicit` による変数宣言の強制

---

# VBARunner テスト作成ガイド

## 目次

1. [セットアップ手順](#セットアップ手順)
2. [テストの書き方](#テストの書き方)
3. [Excel オブジェクトのモック設計](#excel-オブジェクトのモック設計)
4. [前提知識](#前提知識)
5. [バグレポート](#バグレポート)
6. [既知の制限事項](#既知の制限事項)

---

## セットアップ手順

### 1. パッケージの準備

vba-runner は npm パッケージとして配布される。
このワークスペースには VS Codeからの拡張機能のデバッグ実行([VSCODE_SAMPLE_WORKSPACE.md](../../docs/VSCODE_SAMPLE_WORKSPACE.md)参照)した場合は、以下のインストールは完了した状態になっている。

```bash
cd sample/workspace
npm install
```

`package.json` に以下が設定されていれば OK:

```json
{
  "dependencies": {
    "vba-runner": "file:../vba-runner-0.1.0.tgz"
  }
}
```

### 2. テストファイルの作成

`.cjs` 形式（CommonJS）で作成すれば TypeScript コンパイル不要で Node.js から直接実行できる。

```bash
node DataProcessor.test.cjs
```
---

## テストの書き方

### 基本構造

```js
'use strict';

const { VBARunner, assert } = require('vba-runner');

async function main() {
    const runner = new VBARunner('./MyModule.bas');

    // Excel ビルトイン関数のスタブ（必要に応じて）
    runner.set('rgb', (r, g, b) => r + g * 256 + b * 65536);

    // テスト
    assert.strictEqual(runner.run('MyFunction', [arg1, arg2]), expected, 'メッセージ');
}

main().catch(e => { console.error(e); process.exit(1); });
```

### VBARunner の主要 API

| API | 説明 |
|---|---|
| `new VBARunner(path)` | `.bas` / `.cls` ファイル、またはディレクトリを渡す |
| `runner.run(name, args)` | 名前付き Sub/Function を引数付きで呼び出す |
| `runner.eval(expr)` | VBA の式・文を文字列として評価する |
| `runner.set(name, value)` | VBA 環境に変数・関数を注入する |
| `runner.setConstants(obj)` | 複数の定数を一括注入する |
| `runner.evaluator` | 内部 `Evaluator` インスタンスへの直接アクセス |

### パスの解決

`new VBARunner('./DataProcessor.bas')` のパスは **Node.js プロセスの実行ディレクトリ（cwd）** 基準で解決される。`sample/workspace/` から実行するなら `./DataProcessor.bas` でよい。

---

## Excel オブジェクトのモック設計

Excel に依存する VBA コードをテストするには JavaScript でワークシートを模倣するオブジェクトを作る必要がある。

### 識別子の正規化（重要）

VBA エンジンは **識別子をすべて小文字に正規化** する。モックのプロパティ名はすべて小文字で定義しなければならない。

| VBA コード | モックのキー |
|---|---|
| `ws.Cells(r, c)` | `ws.cells` |
| `ws.Rows.Count` | `ws.rows.count` |
| `ws.Columns.AutoFit` | `ws.columns.autofit` |

### 「関数かつプロパティ」な Excel オブジェクト

`ws.Rows(1)` と `ws.Rows.Count` のように、**呼び出し可能でありかつプロパティを持つ**オブジェクトが多い。JavaScript の関数はオブジェクトなので、関数にプロパティを付ける方法で実現する。

```js
ws.rows = function(r) {
    return { copy: ..., font: ..., interior: ... };
};
ws.rows.count = 1048576;  // ws.Rows.Count
```

同様に `ws.cells.clear()` も:

```js
ws.cells = function(r, c) { return { ... }; };
ws.cells.clear = function() { ... };
```

### 行コピー（`Rows(i).Copy destCell`）の実装

`wsSource.Rows(i).Copy wsDest.Cells(destRow, 1)` のコピー先は**セルオブジェクト**として渡される。コピー先セルにシート参照（`_ws`, `_r`）を持たせると行コピーを実装できる。

```js
ws.cells = function(r, c) {
    return {
        _ws: ws, _r: r, _c: c,
        get value() { return cellData[`${r},${c}`] ?? ''; },
        set value(v) { cellData[`${r},${c}`] = v; },
        ...
    };
};

ws.rows = function(r) {
    return {
        copy: function(destCell) {
            const destWs  = destCell._ws;
            const destRow = destCell._r;
            for (let c = 1; c <= 10; c++) {
                const val = cellData[`${r},${c}`];
                if (val !== undefined) destWs._data[`${destRow},${c}`] = val;
            }
        },
        ...
    };
};
```

### `End(xlUp).Row`（最終行取得）の実装

```vba
LastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
```

`xlUp` は Excel 定数（値 `-4162`）。エンジンでは未定義なので `0` として扱われる。方向引数は無視してカラムを線形スキャンする実装で十分。

```js
end: function(_dir) {
    let lastRow = 1;
    for (let row = 1; row <= 10000; row++) {
        const v = cellData[`${row},${c}`];
        if (v !== undefined && v !== '') lastRow = row;
    }
    return { row: lastRow };
}
```

---

## 前提知識

### Option Explicit チェック

VBA エンジンは `Option Explicit` が有効なモジュールに対して、**プロシージャ呼び出し直前**に未宣言識別子チェックを行う。チェックは AST を静的に走査するが、**エラー判定時点**での実行環境（env）を参照するため、`runner.set()` や別モジュールの定義で解決できる。

#### Excel 型ライブラリ定数（`xlUp` 等）の扱い

Excel VBA では `xlUp`・`xlDown`・`vbCrLf`・`vbInformation` などは Excel/VBA 型ライブラリが自動で提供する。当エンジンには型ライブラリがないため、使用する定数を明示的に注入する必要がある。

```js
// Option Explicit が有効な .bas で xlUp を使う場合
runner.set('xlup', -4162);          // Excel.XlDirection.xlUp
runner.set('xldown', -4121);        // Excel.XlDirection.xlDown
runner.set('xltoright', -4161);
runner.set('xltoleft', -4159);
runner.set('vbinformation', 64);    // VBA 標準定数（登録済みが多いが念のため）
```

`runner.set()` は `runner.run()` より前に呼べばよい（ロード順は問わない）。

複数の定数をまとめて注入する場合は `setConstants()` が使いやすい：

```js
runner.setConstants({
    xlup:          -4162,
    xldown:        -4121,
    xltoright:     -4161,
    xltoleft:      -4159,
    vbinformation: 64,
});
```

別の `.bas` ファイルに `Public Const xlUp As Long = -4162` として定義し、`VBARunner` コンストラクターにディレクトリを渡してまとめてロードする方法でも動作する。

### Excel ビルトイン関数のスタブ

`RGB`・`MsgBox` など Excel / VBA の組み込み関数の一部はエンジンに未実装。`runner.set()` で注入できる。

```js
runner.set('rgb',    (r, g, b) => r + g * 256 + b * 65536);
runner.set('msgbox', (msg) => console.log('[MsgBox]', msg));
```

### Private プロシージャのテスト

`runner.run()` は `Private` な Sub/Function も名前で直接呼び出せる。アクセス制御はテスト環境では適用されない。
