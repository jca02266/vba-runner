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

vba-runner は npm パッケージとして配布される。`file:` 参照・`npm link` ではなく **`npm pack` してインストール** することで publish 後と同じ状態を再現できる。

```bash
# モノレポルートで実行
npm run build:lib    # dist/lib.cjs を生成
npm run build:bin    # dist/bin/vba-*.cjs を生成
npm pack             # vba-runner-x.x.x.tgz を生成
```

### 2. テストワークスペースにインストール

```bash
cd vba-sample
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

### 3. テストファイルの作成

`.cjs` 形式（CommonJS）で作成すれば TypeScript コンパイル不要で Node.js から直接実行できる。

```bash
node DataProcessor.test.cjs
```

または `package.json` にスクリプトを追加してから:

```json
{
  "scripts": {
    "test": "node DataProcessor.test.cjs"
  }
}
```

```bash
npm test
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

`new VBARunner('./DataProcessor.bas')` のパスは **Node.js プロセスの実行ディレクトリ（cwd）** 基準で解決される。`vba-sample/` から実行するなら `./DataProcessor.bas` でよい。

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

---

## バグレポート

### BUG-001: For ループ内ラベルへの GoTo がラベル未発見エラーになる

#### 概要

| 項目 | 内容 |
|---|---|
| **重大度** | High（VBA 正規コードが実行不能） |
| **影響範囲** | ループ本体内の `GoTo label` + 同ループ内の `label:` の組み合わせ |
| **対象ファイル** | `src/engine/evaluator.ts` — `executeStatements` / `evaluateForStatement` |
| **発見日** | 2026-05-30 |
| **発見状況** | DataProcessor.bas のテスト作成中 |

#### 再現コード

```vba
Option Explicit

Sub TestGoToInFor()
    Dim i As Long
    Dim seenIDs As Object
    Set seenIDs = CreateObject("Scripting.Dictionary")

    For i = 1 To 5
        If seenIDs.exists(i) Then GoTo SkipRow   ' ← 2 回目以降に GoTo が実行される
        seenIDs.Add i, True
        seenIDs.Add i, True                       ' 2 回目で重複 → GoTo が実行される
SkipRow:
    Next i
End Sub
```

#### エラーメッセージ

```
Error: Sub or Function not defined: label 'SkipRow' (line XX)
  type: 'VbaError', number: 35
```

#### 期待動作

実際の Excel VBA では、`GoTo` のジャンプ先ラベルは**手続き内のどの位置にあってもよい**（ループ内・ネスト内を問わない）。MS-VBAL §5.2.1.2 でラベルのスコープは手続き（procedure）であると定義されている。

#### 根本原因

`executeStatements` のラベル解決は **「呼び出し元の `body` 配列を `findIndex` で線形検索する」** 設計になっている。

```
手続き executeStatements(procBody, isTopLevel=true)
  └─ ForStatement                           ← procBody に含まれる
       └─ executeStatements(forBody, isTopLevel=false)
            └─ IfStatement (inline)
                 └─ executeStatements(ifBody, isTopLevel=false)
                      └─ GoToStatement → throw { type: "GoTo", label: "SkipRow" }
                 ↑ isTopLevel=false → re-throw
            ↑ isTopLevel=false → re-throw (ForStatement は "Exit" だけキャッチ)
  ↑ isTopLevel=true → body.findIndex で procBody を検索
       → "SkipRow:" は forBody 内にあるため procBody には存在しない → -1 → Error
```

問題の核心は **`isTopLevel=false` の `executeStatements` が GoTo 例外を即 re-throw し、`forBody` の中にあるラベルを自分では探さない** 点にある。

```javascript
// evaluator.ts (簡略)
executeStatements(body, startIndex, isTopLevel = true) {
    // ...
    } catch (e) {
        // isTopLevel=false の場合、GoTo を無条件で上に投げ返す ← 問題箇所
        if (!isTopLevel && e?.type === "GoTo") {
            throw e;
        }
        // isTopLevel=true の場合のみラベル解決を試みる
        if (e?.type === "GoTo") {
            const labelIndex = body.findIndex(  // ← この body は手続きトップレベル
                s => s.type === "LabelStatement" && s.label === e.label
            );
            if (labelIndex < 0) throwVbaError("Sub or Function not defined: label ...");
        }
    }
}

evaluateForStatement(stmt) {
    while (condition()) {
        try {
            this.executeStatements(stmt.body, 0, false);  // false = isTopLevel
        } catch (e) {
            if (e?.type === "Exit" && e.target === "For") break;
            throw e;  // GoTo もここで re-throw される ← For ループ自身もラベルを探さない
        }
    }
}
```

#### 修正方針（提案）

`evaluateForStatement`（および `evaluateDoWhileStatement`, `evaluateWhileStatement`）で、re-throw 前に自身の `stmt.body` でラベルを探し、見つかれば GoTo を処理する。

```javascript
// 修正案
evaluateForStatement(stmt) {
    while (condition()) {
        try {
            this.executeStatements(stmt.body, 0, false);
        } catch (e) {
            if (e?.type === "Exit" && e.target === "For") break;

            // ← 追加: ループ内のラベルへの GoTo をここで解決する
            if (e?.type === "GoTo") {
                const labelIndex = stmt.body.findIndex(
                    s => s.type === "LabelStatement"
                      && s.label.toLowerCase() === e.label.toLowerCase()
                );
                if (labelIndex >= 0) {
                    // ラベル位置から再開（executeStatements を分割実行）
                    this.executeStatements(stmt.body, labelIndex, false);
                    continue;
                }
            }
            throw e;
        }
        this.env.setLocally(varName, this.env.get(varName) + stepValue);
    }
}
```

同様の対処が `evaluateDoWhileStatement`, `evaluateWhileStatement`, `evaluateWithStatement` にも必要。

#### 隠れた問題：テストが偽陽性で通過する

`CopyFilteredData` も For ループ内で `GoTo NextRow` を使っているが、テストデータに空 ID の行がなく GoTo が一度も実行されなかったため通過した。

```vba
' CopyFilteredData 内 — GoTo が実行されると同じバグが顕在化する
For i = 2 To lastRow
    If currentID = "" Then GoTo NextRow  ← 今回のテストでは実行されなかった
    ...
NextRow:
Next i
```

**GoTo を含む手続きをテストするときは、GoTo が実際に実行される入力データを必ず含めること。**

#### 回避策（現時点）

`GoTo` を使わず `If/End If` のネストで書き換える。

```vba
' Before（動かない）
For i = 2 To lastRow
    If seenIDs.exists(currentID) Then GoTo SkipRow
    ' ... 処理 ...
SkipRow:
Next i

' After（回避策）
For i = 2 To lastRow
    If Not seenIDs.exists(currentID) Then
        ' ... 処理 ...
    End If
Next i
```

---

## 既知の制限事項

### LIMIT-001: Option Explicit + 未宣言 Excel 定数

Excel 型ライブラリの定数（`xlUp = -4162`, `vbCrLf`, `vbInformation` 等）は、型ライブラリ参照なしの環境では宣言なしで使われているため `Option Explicit` チェックで違反とみなされ、そのプロシージャが呼び出し不能になる。

**チェックのタイミング**: ランタイムではなく `.bas` ロード時（`evaluate` フェーズ）の静的解析。チェック後に `runner.set('xlup', ...)` で注入しても手遅れ。

**対処**: `.bas` ファイルに定数宣言を追加する。

```vba
Private Const xlUp        As Long = -4162
Private Const vbCrLf      As String = Chr(13) & Chr(10)
Private Const vbInformation As Long = 64
```

---

### LIMIT-002: RGB 関数が未実装

`RGB(r, g, b)` は VBA 標準ライブラリ関数だが 2026-05 時点で未実装。`runner.set` でスタブを注入して回避する。

```js
runner.set('rgb', (r, g, b) => r + g * 256 + b * 65536);
```

正確な計算式: `RGB(r, g, b) = r | (g << 8) | (b << 16)`（Excel の OLE_COLOR は COLORREF 形式）。

---

### LIMIT-003: Scripting.Dictionary のキー列挙順

#### エンジンの実装

内部実装は JavaScript の `Map`。`Map` はキーを**挿入順**で列挙する（ECMAScript 仕様で保証）。

```javascript
// evaluator.ts 内
this.registerExternalObject("Scripting.Dictionary", () => {
    const dict = new Map();
    return {
        keys:  () => Array.from(dict.keys()),   // 挿入順
        items: () => Array.from(dict.values()),  // 挿入順
        // ...
    };
});
```

#### 実際の VBA（Excel）の挙動

Microsoft の `Scripting.Dictionary`（`scrrun.dll`）も**挿入順**でキーを返す。これは文書化された動作ではないが、内部実装（ハッシュテーブル + 挿入順リンクリスト）から来る実質的な仕様として広く使われており、VBA コードは挿入順を前提に書かれることが多い。

#### 結論

エンジンの挿入順動作は実際の VBA と一致しており、このテストでは問題にならない。ただし、順序依存のアサーション（「カテゴリ A が行 5、カテゴリ B が行 6」など）はデータ投入順に依存するため、テストコードにその前提を明示しておくこと。

---

## 今回変更したファイル

| ファイル | 変更内容 |
|---|---|
| `DataProcessor.bas` | `Private Const xlUp` 追加 / `SummarizeByCategory` の GoTo を If/End If に書き換え（BUG-001 回避） |
| `DataProcessor.test.cjs` | テストファイル（新規作成） |
| `package.json` | `"test"` スクリプト追加 |
