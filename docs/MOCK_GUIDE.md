# VBA モック実装ガイド

> **まずは Part 0 を試してください**: `__mocks__/` ディレクトリ(または `__mocks__.*` 単一ファイル)を
> 置くだけでビルトインを差し替えられる規約で、npm パッケージ利用者・VS Code 拡張機能利用者のどちらでも
> リポジトリ内部のソースパスを import せずに使えます。`Application.OnKey` / `OnTime` のように
> npm の公開 API(`setBuiltinOverride` 等)でも差し替えにくいケースも、`__mocks__/` の VBA `Class` で
> 上書きできます。
>
> **Part 1 以降のコード例は一部npmパッケージから直接使えません**: `MockApplication`/`MockWorksheet`/
> `MockRange`/`MockRows`/`MockColumns`/`MockWorkbook`自体は`vba-runner`からexportされており
> `import { MockApplication } from 'vba-runner'`で使えるが、`import { Lexer } from '../../src/engine/lexer'`や
> `new Evaluator(...)`を直接構築するような低レベルAPIはリポジトリ内部のソースパス前提のままで、
> npmでインストールしたパッケージから直接importできない。npm利用者はこれらの代わりに
> `VBARunner`(`vbaRunner.evaluator`で内部の`Evaluator`インスタンスにもアクセスできる)を使う。
> Part 0 だけでは足りない場合、npm経由で同等のことを行うには以下の公開APIを使ってください:
>
> - `new VBARunner(path, { excelStub: true | MockApplicationのインスタンス })` →
>   `runner.excelStub.ActiveSheet.setCellValue(...)` / `getCellValue(...)` でセルの初期値設定・検証。
>   `excelStub`に`MockApplication`をサブクラス化したインスタンスを渡せば、型安全に拡張できる
>   (`OnKey`/`OnTime`の例は本ファイルの「現在対応していない機能」直後の注記参照)
> - `runner.set(name, value)` で任意の名前のグローバル変数・定数を注入
> - `runner.evaluator.setBuiltinOverride(name, value)` で `Application` 等の組み込みオブジェクトを
>   丸ごと独自オブジェクトに差し替え(`MockApplication` を継承する代わりに、必要なメソッドだけを
>   持つ素のオブジェクトを自作して渡せばよい。ただし置き換えではなく拡張で済む場合はそちらを優先すること)
>
> 詳しくは [`README.md`](https://github.com/jca02266/vba-runner/blob/main/build/runner/README.md)
> の「Mocking Excel-dependent objects」を参照してください。
> Part 1 以降の低レベルAPIの内容(リポジトリ内部のソースパスを使う例)は、vba-runner本体の開発に
> 参加する場合や、リポジトリをcloneして拡張する場合に有用です。

---

# 📍 Part 0: `__mocks__` 規約で差し替える（推奨、まずはここを試す）

VBA コードが `MsgBox` や `Range` といった Excel / Windows 環境に依存する処理を呼び出すとき、
エンジン単体での実行はエラーになります。`__mocks__/` ディレクトリ（または `__mocks__.*` ファイル）に
モックを置くと、`VBARunner` / 拡張機能の `▶ Run` 実行時にエンジンのビルトインが自動的に差し替えられ、
Excel がなくても VBA コードを実行できるようになります。

> **環境依存処理だけが対象ではない**: `__mocks__/` は「同名の関数・クラスを補う/上書きする」
> という汎用的な仕組みなので、`RGB` のように Excel/Windows に依存しない**単純に未実装の組み込み
> 関数**を補う用途にもそのまま使える(→ [0-5](#0-5-未実装の組み込み関数を補うrgbの例))。
> エンジンが対応していない組み込み関数に遭遇したら、まず`__mocks__/`で補えないか検討するとよい。

## ひな形の自動生成（vba-runner.generateMocks）

Excel 依存のコードに対して `__mocks__/ExcelObjects.bas` のひな形を自動生成できます。

**VS Code から使う:**

1. Excel 依存の `.bas` ファイルをエディターで開く
2. コマンドパレット（Cmd+Shift+P）→ `VBA: Generate Mock Stubs` を実行
3. `__mocks__/ExcelObjects.bas` が生成され、自動的に開く

**生成されるファイルの例（`ActiveSheet`, `Range`, `Cells`, `Rows` が検出された場合）:**

```vba
' __mocks__/ExcelObjects.bas
' 自動生成: vba-runner.generateMocks 2026-06-07
' 必要に応じてセル値・戻り値を書き換えてください

Class MockRange
    Public Value As Variant   ' ← 読み取り値をここで設定
    Public Row As Long
    Public Column As Long
    ' ※ Range.End(xlUp) のような予約語名プロパティは .bas スタブでは定義できません。
    '   VBA の仕様上、プロシージャ名には IDENTIFIER（予約語を除く）しか使えないためです。
    '   仕様上の回避策は [End] という FOREIGN-NAME 構文ですが、現エンジンでは未実装です。
    '   → __mocks__.js / __mocks__.ts で JS クラスとして定義してください。
    '     JS では End は予約語でなくプロパティ名に使え、VBA 側からの r.End(xlUp) 呼び出しも動作します。
End Class

Class MockRows
    Public Count As Long
    Sub Class_Initialize()
        Count = 1048576
    End Sub
End Class

Class MockWorksheet
    Public Property Get Name() As String
        Name = "Sheet1"
    End Property

    Public Function Cells(row, col)
        Dim r As New MockRange
        r.Row = row
        r.Column = col
        ' r.Value = 0  ← テスト用の値はここに設定
        Set Cells = r
    End Function

    Public Function Range(addr)
        Dim r As New MockRange
        ' r.Value = 0  ← テスト用の値はここに設定
        Set Range = r
    End Function

    Public Property Get Rows()
        Set Rows = New MockRows
    End Property
End Class

Function ActiveSheet()
    Set ActiveSheet = New MockWorksheet
End Function
```

**カスタマイズ方法:**

生成後にセル値を固定したい場合は `Cells` / `Range` 関数内の `r.Value` に値を設定します。

```vba
Public Function Cells(row, col)
    Dim r As New MockRange
    r.Row = row
    r.Column = col
    If row = 1 And col = 1 Then r.Value = 100  ' ← A1 = 100 として動作する
    If row = 2 And col = 1 Then r.Value = 200  ' ← A2 = 200 として動作する
    Set Cells = r
End Function
```

**`injectExcelStub` との関係:**

`▶ Run` 実行時は常に `injectExcelStub`（サイレント注入）が先に適用されます。
その後 `__mocks__/ExcelObjects.bas` がロードされ、同名の関数・クラスを上書きします。
つまり **生成ファイルに書いた定義だけが選択的に上書きされ**、書いていない部分はサイレント注入のままです。

***

## モック形式の選び方

モックは `.bas`（VBA）・`.js`（JavaScript）・`.ts`（TypeScript）の 3 形式で書けます。
どれを選ぶかは「何をモックしたいか」と「どの環境で動かすか」で決まります。

| 形式     | 追加インストール              | JS ライブラリ利用               | 向いている用途                           |
| ------ | --------------------- | ------------------------ | --------------------------------- |
| `.bas` | 不要                    | **不可**（VBA の範囲内のみ）       | Excel オブジェクトの構造スタブ・単純なロジック差し替え    |
| `.js`  | 不要                    | 可（Node.js API・npm パッケージ） | COM 依存の代替・正規表現・HTTP など JS ネイティブ機能 |
| `.ts`  | tsx 等が必要(VS Code は任意) | 可（型安全）                   | `.js` と同じことを型チェック付きで書きたい場合        |

`.bas` は VBA エンジン上で評価されるため、`require()` や `import` は使えません。
正規表現・HTTP クライアント・ファイルシステムなど Node.js ネイティブの機能が必要なモックは
`.js` または `.ts` で書く必要があります。

`.ts` モックのロード方法は実行環境によって異なります。CLI(vba-runner npmパッケージ)から `npx tsx` 経由でテストを実行している場合は
tsx がモジュールローダーに TypeScript 変換フックを登録するため、ビルドなしで `require()` できます。
VS Code 拡張機能から実行する場合は tsx 等の TypeScript 実行拡張機能が別途インストールされていれば
同様に動きます（追加インストールが必要）。`.js` か `.bas` を使う場合はいずれの環境でも追加インストールなしで動かすことができます。

## ロード順と優先順位

モックファイルは **ファイル名（basename）の ASCII 辞書順にロードされ、後にロードされたファイルが先のファイルを上書きします（後勝ち）。**

ソートキーは basename のみで、`__mocks__/` ディレクトリ内のファイルならそのファイル名、`__mocks__.*` 単一ファイルなら `__mocks__.<ext>` がソートキーになります。

ASCII コードの大小関係: 大文字 A–Z (65–90) < アンダースコア _ (95) < 小文字 a–z (97–122)

大文字始まりのファイル名（`ExcelObjects.bas` 等）は `__mocks__.*` より ASCII 値が小さいため **先にロードされ、優先度が低くなります。**

典型的なソート順と優先度の例（低 → 高）:

| ロード順 | ファイル（ソートキー） | 優先度 |
|---|---|---|
| 1 | `ExcelObjects.bas` (E=69) | 最低 |
| 2 | `MsgBox.js` (M=77) | ↑ |
| 3 | `RegExp.ts` (R=82) | ↑ |
| 4 | `__mocks__.bas` (_=95, b=98) | ↑ |
| 5 | `__mocks__.js` (_=95, j=106) | ↑ |
| 6 | `__mocks__.ts` (_=95, t=116) | 最高 |

同じ関数名やクラス名が複数のファイルに定義されている場合、ソート順で最後のファイルの定義が使われます。

**`__mocks__.<ext>` 単一ファイルは常に `__mocks__/` ディレクトリ内の同拡張子ファイルより後にロードされるため、優先度が高くなります。**（`_` > 大文字という ASCII の大小関係による）

同拡張子の単一ファイル間では `.bas` < `.js` < `.ts` の順なので、`.ts` が最高優先です。

***

## 0-1. MsgBox をコンソール出力に差し替える

VBA の `MsgBox` はダイアログを出してユーザーの応答を待つため、
自動テストや CI 環境では実行がブロックされてしまいます。
モックで `Debug.Print` や `console.log` に差し替えることで、止まらずに実行できます。

3 つの形式すべてで同じことができます。

### .bas で書く場合

追加インストール不要で、VBA の `Debug.Print` に転送します。

```
src/
├── Main.bas
└── __mocks__.bas
```

```vba
' __mocks__.bas
Function MsgBox(prompt, Optional buttons, Optional title)
    Debug.Print "[MsgBox] " & IIf(title <> "", title & ": ", "") & prompt
    MsgBox = 1  ' vbOK
End Function

Function InputBox(prompt, Optional title, Optional defaultValue)
    Debug.Print "[InputBox] " & IIf(title <> "", title & ": ", "") & prompt
    InputBox = IIf(IsMissing(defaultValue), "", defaultValue)
End Function
```

### .js で書く場合

Node.js の `console.log` に転送します。追加インストール不要です。

```
src/
├── Main.bas
└── __mocks__.js
```

```javascript
// __mocks__.js
module.exports = {
    MsgBox: (prompt, _buttons, title) => {
        console.log(`[MsgBox] ${title ? title + ': ' : ''}${prompt}`);
        return 1; // vbOK
    },
    InputBox: (prompt, title, defaultValue) => {
        console.log(`[InputBox] ${title ? title + ': ' : ''}${prompt}`);
        return defaultValue ?? '';
    },
};
```

### .ts で書く場合

`.js` と同じですが、引数の型が明示されるため取り違えをコンパイル時に検出できます。
`npx tsx` 環境ではビルド不要です。

```typescript
// __mocks__.ts
export default {
    MsgBox: (prompt: string, _buttons?: number, title?: string): number => {
        console.log(`[MsgBox] ${title ? title + ': ' : ''}${prompt}`);
        return 1; // vbOK
    },
    InputBox: (prompt: string, title?: string, defaultValue?: string): string => {
        console.log(`[InputBox] ${title ? title + ': ' : ''}${prompt}`);
        return defaultValue ?? '';
    },
};
```

***

## 0-2. VBA クラスで Excel オブジェクトをスタブする

`__mocks__/` に置いた VBA クラスはエンジンの `externalObjectFactories` に登録されるため、
`Set ws = ActiveSheet` のようなコードをエンジン上で動かせるようになります。

ここでは `Worksheet` を例に説明します。`Worksheet` はグローバル関数として呼ばれることがないため、
クラス名と関数名の衝突が起きない、VBA クラスモックに向いているオブジェクトです。

__`mocks`__ のソースをExcel等本物のVBAで動かす必要はないので、本VBAエンジンの拡張構文であるClassを使用しています。これにより、クラスごとに複数ファイルのソースを作る手間を省いています。

```
src/
├── CalcTotal.bas
└── __mocks__/
    └── ExcelObjects.bas
```

```vba
' __mocks__/ExcelObjects.bas

Class Worksheet
    Private mCells(100, 100) As Variant

    Public Function Cells(row, col)
        Set Cells = New CellRef
        Cells.SetRef mCells, row, col
    End Function

    Public Property Get Name() As String
        Name = "MockSheet"
    End Property
End Class

Class CellRef
    Private mStorage() As Variant
    Private mRow As Long
    Private mCol As Long

    Public Sub SetRef(storage, row, col)
        mStorage = storage
        mRow = row
        mCol = col
    End Sub

    Public Property Get Value() As Variant
        Value = mStorage(mRow, mCol)
    End Property

    Public Property Let Value(v As Variant)
        mStorage(mRow, mCol) = v
    End Property
End Class

Function ActiveSheet()
    Set ActiveSheet = New Worksheet
End Function

Function Sheets(nameOrIndex)
    Set Sheets = New Worksheet
End Function
```

これで VBA コード側はそのまま動きます。

```vba
Dim ws As Worksheet
Set ws = ActiveSheet
ws.Cells(1, 1).Value = 100
Debug.Print ws.Cells(1, 1).Value  ' => 100
```

### 制約：エンジンがデフォルトオブジェクト経由の名前解決を未実装

`Range` を例に説明します。実 VBA では `Range("A1")` は `ActiveSheet.Range.Item("A1")` の糖衣構文であり、
`Range` はグローバル関数ではなく `Worksheet` のプロパティです。また `Dim r As Range` の `Range` は
型名前空間に属し、プロパティアクセスの名前空間とは別物なので、実 VBA では名前の衝突は起きません。

現状のエンジンは `Range("A1")` をグローバル関数呼び出しとして解決しており、
「デフォルトオブジェクト（`ActiveSheet`）経由のプロパティアクセス」は未実装です。

**部分改善（§5.6.10 Tier 6 対応）**: クラスモジュール名が値名前空間に登録されることで
`Range("A1")` の呼び出しが `OBJECT_REQUIRED` で失敗していた問題を修正しました。
現在は以下のパターンが可能です:

```javascript
// TypeScript テストコード
const ev = evalVBAModules([
    { name: 'Range', code: rangeMockCode },   // Class Range の定義
    { name: 'UserModule', code: userCode }
]);
// Range ファクトリ関数を後から登録
ev.set('Range', (addr) => {
    const r = ev.instantiateClass('Range');
    r.__instanceEnv__.set('value', addr);
    return r;
});
```

または `.bas` ファイル内で `Class Range` と `Function Range` を同居させる方法もあります:

```vba
' ExcelMocks.bas（モジュール名は 'ExcelMocks' など、'Range' 以外）
Class Range
    Public Value As Variant
End Class

Function Range(address As String) As Range
    Dim r As New Range
    r.Value = address
    Set Range = r
End Function
```

より複雑な用途には `.js` / `.ts` モックでのビルトイン拡張を推奨します
（→ [0-3. ビルトインをラップする](#0-3-ビルトインをラップする構想)）。

***

## 0-3. ビルトインをラップする（構想）

> **現時点では未実装です。将来の継承機構として検討中の設計です。**

現状、`__mocks__/` でビルトインを差し替える場合はメソッドをすべて再実装する必要があります。
例えば `Dictionary` に操作ログを追加したいだけでも、`Add` / `Item` / `Remove` 等をすべて
自分で書き直さなければなりません。

将来的には、`__wrap__: true` を指定すると元のビルトイン実装を受け取りながら挙動だけを
上書きできる「ラップ」機構を導入したいと考えています。
これにより、差し替えたいメソッドだけを書けばよくなります。

元のビルトインへの参照は `__mock_super__` という名前で各メソッドの第1引数として渡されます。
`super` に似た命名にすることで「元の実装を呼び出す継承的な操作」であることが伝わります。

```javascript
// __mocks__/Dictionary.js（構想）
module.exports = {
    Dictionary: {
        __wrap__: true,
        // __mock_super__ にオリジナルのビルトインメソッドが渡ってくる
        Add: (__mock_super__, key, value) => {
            console.log(`[Dictionary.Add] "${key}" = ${value}`);
            return __mock_super__(key, value); // 元の Add を呼ぶ
        },
        Item: (__mock_super__, key) => {
            const result = __mock_super__(key); // 元の Item を呼ぶ
            console.log(`[Dictionary.Item] "${key}" => ${result}`);
            return result;
        },
    },
};
```

***

## 0-4. 正規表現ライブラリのモック（JS モックが必要な例）

VBA で正規表現を使うには `CreateObject("VBScript.RegExp")` で COM オブジェクトを生成しますが、
エンジンは COM を呼び出せないためそのままではエラーになります。
これは `.bas` モックでは絶対に解決できない問題です。VBA 自体に正規表現エンジンがなく、
COM に完全依存しているからです。

`.js` / `.ts` モックであれば Node.js ネイティブの `RegExp` をそのまま使って代替できます。

`CreateObject` 用のファクトリは通常のビルトイン上書き（`MsgBox` 等）と区別するため、
`__addCreateObject__` キーの下に ProgID をキーとして登録します。

エンジンは `__createObject__` という内部レジストリで全 ProgID ファクトリを管理しており、
ビルトインも同じレジストリに登録されています。モックファイルが `__addCreateObject__` を export すると、
ローダーがその中身を `__createObject__` レジストリに追記します。
`__addCreateObject__` というキー名にすることで、複数のモックファイルが同じキーを export しても
ローダーが順番にマージするだけで衝突が起きません。

```javascript
// __mocks__/RegExp.js
class VBScriptRegExp {
    constructor() {
        this.Pattern = '';
        this.Global = false;
        this.IgnoreCase = false;
        this.MultiLine = false;
    }

    _flags() {
        return (this.Global ? 'g' : '') +
               (this.IgnoreCase ? 'i' : '') +
               (this.MultiLine ? 'm' : '');
    }

    Test(str) {
        return new RegExp(this.Pattern, this._flags()).test(String(str));
    }

    Replace(str, replacement) {
        return String(str).replace(new RegExp(this.Pattern, this._flags()), replacement);
    }

    Execute(str) {
        const matches = [...String(str).matchAll(new RegExp(this.Pattern, this._flags() + 'g'))];
        return {
            Count: matches.length,
            Item: (i) => ({ Value: matches[i]?.[0] ?? '' }),
        };
    }
}

module.exports = {
    // __addCreateObject__ の中身がエンジンの __createObject__ レジストリに追記される
    __addCreateObject__: {
        'VBScript.RegExp': () => new VBScriptRegExp(),
    },
};
```

VBA コード側は変更不要です。

```vba
Dim re As Object
Set re = CreateObject("VBScript.RegExp")
re.Pattern = "\d+"
re.Global = True
Debug.Print re.Test("abc123")          ' => True
Debug.Print re.Replace("abc123def456", "#")  ' => "abc#def#"
```

TypeScript で書く場合はインターフェースを型で表現でき、
メソッドの実装漏れをコンパイル時に検出できます。

```typescript
// __mocks__/RegExp.ts
interface MatchCollection {
    Count: number;
    Item: (i: number) => { Value: string };
}

class VBScriptRegExp {
    Pattern = '';
    Global = false;
    IgnoreCase = false;
    MultiLine = false;

    private flags(): string {
        return (this.Global ? 'g' : '') +
               (this.IgnoreCase ? 'i' : '') +
               (this.MultiLine ? 'm' : '');
    }

    Test(str: string): boolean {
        return new RegExp(this.Pattern, this.flags()).test(str);
    }

    Replace(str: string, replacement: string): string {
        return str.replace(new RegExp(this.Pattern, this.flags()), replacement);
    }

    Execute(str: string): MatchCollection {
        const matches = [...str.matchAll(new RegExp(this.Pattern, this.flags() + 'g'))];
        return {
            Count: matches.length,
            Item: (i: number) => ({ Value: matches[i]?.[0] ?? '' }),
        };
    }
}

export default {
    __addCreateObject__: {
        'VBScript.RegExp': () => new VBScriptRegExp(),
    },
};
```

***

## 0-5. 未実装の組み込み関数を補う(RGBの例)

`RGB`のように、Excel/Windows環境に依存しない**単純にエンジンへの実装が漏れている組み込み関数**も、
`__mocks__/`で同名の関数を定義するだけで補える(Excel依存の差し替えとまったく同じ仕組み)。

```vb
' __mocks__/Builtins.bas
Function RGB(r As Long, g As Long, b As Long) As Long
    RGB = r + g * 256 + b * 65536
End Function
```

これで呼び出し側のVBAコードは実Excelと同じ`RGB(r, g, b)`をそのまま書ける。独自名のヘルパー関数
(`MakeColor`等)をソースに混ぜる必要がない。`vba-extractor import`はサブディレクトリを読まないため、
このスタブが実際の`.xlsm`に紛れ込むこともない。

> **テストコード内だけで済ませたい場合**: `__mocks__/`にファイルを置く代わりに、
> `runner.set('rgb', (r, g, b) => r + g * 256 + b * 65536)`のように`VBARunner.set()`で
> その場で注入することもできる(プロパティ名は小文字)。`▶ Run`実行や複数テストファイルで
> 共有したい場合は`__mocks__/`、個々のテストファイル内だけで一時的に注入したい場合は
> `runner.set()`、と使い分けるとよい。

エンジンが未対応の組み込み関数に遭遇したら(`Sub or Function not defined: 'XXX'`)、まずこの方法で
補えないか検討すること。

***

## JS/TS モックオブジェクトの規約

TypeScript で VBA オブジェクトとして振る舞うクラスを書く際に実装すべきインターフェースの一覧。
すべて `src/engine/vba-types.ts` に定義されており `MockExcel` からも re-export している。

```typescript
import type { VbaType, VbaDefaultProperty, VbaIterable, VbaComObject } from '../src/engine/mock/MockExcel';
```

| インターフェース | マーカー | 用途 | 例 |
|---|---|---|---|
| `VbaType` | `__vbaTypeName__: string` | `TypeOf x Is Range` / `TypeName(x)` | `MockRange`, `MockWorksheet` |
| `VbaDefaultProperty` | `__vbaDefault__: true` + `Value` getter/setter | 読み書き両方で `.Value` を経由する | `MockRange` |
| `VbaIterable` | `[Symbol.iterator]()` | `For Each item In col` | カスタムコレクション |
| `VbaComObject` | `__progId__: string` | `CreateObject(progId)` の自動別名登録 | `MockWordApplication` |

### VbaComObject — CreateObject 自動別名登録

`registerComObject(factory)` は factory() を一度呼び出し、
返り値の `__className__` を読んで同じ factory をその名前でも登録する。
これにより `CreateObject("Word.Application")` と `New Word.Application` の両方が動く。

```typescript
class MockWordApplication implements VbaComObject {
    readonly __progId__ = 'Word.Application';
}
evaluator.registerComObject(() => new MockWordApplication());
// VBA: Set app = CreateObject("Word.Application")  → MockWordApplication
// VBA: Dim app As New Word.Application             → MockWordApplication（自動別名）
```

### VbaType — TypeOf / TypeName 対応

```typescript
class MockRange implements VbaType {
    readonly __vbaTypeName__ = 'Range';  // TypeOf r Is Range → True
    // TypeName(r) → "Range"
}
```

### VbaDefaultProperty — 暗黙 .Value アクセス

evaluator は読み書き両方で `Value` プロパティを経由する（getter/setter 対称）。

```typescript
class MockRange implements VbaDefaultProperty {
    readonly __vbaDefault__ = true as const;
    private _v: any = 0;
    get Value() { return this._v; }
    set Value(v: any) { this._v = v; }
    // VBA: x = Range("A1")   → Value getter の値が x に入る
    // VBA: Range("A1") = 100 → Value setter が呼ばれる
}
```

### VbaIterable — For Each 列挙

```typescript
class MockRows implements VbaIterable {
    private _rows: any[];
    [Symbol.iterator]() { return this._rows[Symbol.iterator](); }
    // VBA: For Each row In Rows → _rows の各要素を列挙
}
```

> **`items` プロパティによる代替**: `[Symbol.iterator]` の代わりに `items: any[]` プロパティでも
> `For Each` に対応できる（エンジンがフォールバック検出する）。ただし `VbaIterable` の型安全は得られない。

### メソッド命名規則

- **PascalCase 推奨**（`Range`, `Cells`, `GetValue`）
- エンジンが大文字小文字を無視して解決するため `range` でも動作するが、
  VBA 側の呼び出しと大文字小文字を合わせておくと見通しがよい

---

## Part 0 の用語について

「モック」という言葉は文脈によって意味が異なります。

狭義の「モック」はテストダブルの一種であり、「どのメソッドが何回呼ばれたか」「どんな引数で呼ばれたか」を事前に期待値として設定し、事後に検証できるオブジェクトを指します。この定義に照らすと、本ドキュメントの例は厳密には「スタブ」（決まった値を返すだけの最小実装）や「フェイク」（本物と同様に動く代替実装）であり、狭義の「モック」の説明としては不適切です。

一方、Jest の `__mocks__/` ディレクトリの慣習に代表されるように、テストダブル全般を広く「モック」と呼ぶ用法も広く普及しています。本ドキュメントはこの広義の「モック」の意味でこの言葉を使っています。

---

## Step 0: モックが必要かどうかを確認する [[→ R-01](REFACTORING_TESTING_CATALOG.md#r-01)]

**モックより先に純粋関数の抽出を試みること。**
Excel オブジェクトへの依存を引数として受け取るように関数を分離できれば、モックは不要になる。

モックが必要になるのは以下のいずれかの場合:
- テスト対象の関数が `ActiveSheet`・`Range`・`Cells`・`Application` 等を直接使っており、引数として切り出せない
- エントリーポイント（マクロ全体）の統合テストを書く必要がある

---

## Step 1: 対応表で注入コードを確認する（ここだけ読めば9割解決）

VBA コード中で使用している Excel オブジェクト名を下表で引く。
該当する「注入コード」を `ev.getGlobalEnv().set(...)` で追加し、詳細が必要なら右列のセクションへ。

| Excel オブジェクト名 | 注入コード | 詳細 |
|---|---|---|
| `ActiveSheet` | `ev.getGlobalEnv().set('ActiveSheet', ws)` | [§A](#a-activesheet--sheets) |
| `Sheets` | `ev.getGlobalEnv().set('Sheets', (n) => mockApp.Sheets(n))` | [§A](#a-activesheet--sheets) |
| `Application`（Sheetsのみ使用）| `ev.getGlobalEnv().set('Application', mockApp)` | [§A](#a-activesheet--sheets) |
| `Application`（ScreenUpdating等）| `ev.getGlobalEnv().set('Application', new MockApplicationWithSettings())` | [§B](#b-application-プロパティscreenupdating-等) |
| `Cells`（`ws.Cells` 経由）| ActiveSheet の注入で解決 | [§A](#a-activesheet--sheets) |
| `Range`（`ws.Range` 経由）| ActiveSheet の注入で解決 | [§A](#a-activesheet--sheets) |
| `Rows`・`Columns`（`ws.Rows.Count` 等）| MockWorksheet が未実装。スタブが必要 | [§C](#c-未対応オブジェクトの拡張) |

> **注意**: `Application.ScreenUpdating` などを注入せずに放置すると、エラーにならず暗黙に無視される場合と、読み取り時に "Type mismatch" で実行が止まる場合がある。詳細は [§B の制限](#b-application-プロパティscreenupdating-等)。

---

## Step 2: 詳細セクション（必要なものだけ読む）

### §A: ActiveSheet / Sheets

最小限の注入パターン:

```typescript
import { MockApplication, MockWorksheet } from '../../src/engine/mock/MockWorksheet';

const mockApp = new MockApplication();
const ws = mockApp.Sheets('Sheet1');

// セル値の設定
ws.setCellValue('A1', 100);
ws.setCellValue('A1:C3', [[1,2,3],[4,5,6],[7,8,9]]);  // 2D配列

ev.getGlobalEnv().set('ActiveSheet', ws);
ev.getGlobalEnv().set('Sheets', (nameOrIndex: string | number) => mockApp.Sheets(nameOrIndex));
ev.getGlobalEnv().set('Application', mockApp);  // Sheets()のみ必要な場合
```

VBA の `ws.Cells(r, c)`・`ws.Range("A1")` は MockWorksheet に実装済みのため、
`ActiveSheet` を注入するだけで自動的に使える。

### §B: Application プロパティ（ScreenUpdating 等）

`MockApplication` は `Sheets()` のみ実装。`ScreenUpdating` などを注入しないと
読み取り時に "Type mismatch" で実行が止まる。小文字プロパティを持つ拡張クラスで対応する:

```typescript
import { MockApplication } from '../../src/engine/mock/MockWorksheet';

class MockApplicationWithSettings extends MockApplication {
    // evaluator は propName.toLowerCase() でアクセスするため小文字で定義
    screenupdating: boolean = true;
    calculation: any = -4105;   // xlCalculationAutomatic 相当
    enableevents: boolean = true;
}

const mockApp = new MockApplicationWithSettings();
ev.getGlobalEnv().set('Application', mockApp);

// 実行後に保存/復元が正しく行われたか確認できる
// assert.strictEqual(mockApp.screenupdating, true, '復元された');
```

> 詳細と仕組みの解説は「§5. MockApplication の制限」セクションを参照。

### §C: 未対応オブジェクトの拡張（Rows.Count / Columns.Count / VBA定数）

**`ws.Rows.Count` / `ws.Columns.Count` のスタブ:**

MockWorksheet はこれらを実装していない。`ActiveSheet` として渡すオブジェクトに直接プロパティを追加する:

```typescript
const ws = mockApp.Sheets('Sheet1') as any;

// Rows.Count / Columns.Count のスタブ
ws.rows = { count: 1000 };   // evaluator は 'rows' (小文字) でアクセス
ws.columns = { count: 256 };

ev.getGlobalEnv().set('ActiveSheet', ws);
```

> evaluator はプロパティ名を小文字化する（`ws.Rows.Count` → `ws['rows']['count']`）。

**`Application.OnKey` / `Application.OnTime` のスタブ:**

`MockApplication` はこれらを実装していない。`Application` として渡しているインスタンス（`excelStub`
で注入されるオブジェクトと同一）に直接メソッドを追加する。**新しいオブジェクトで `Application` を
丸ごと差し替えないこと** — 同じインスタンスを拡張すれば `ActiveSheet`/`Sheets`/`Range` の解決は
影響を受けない。

```typescript
const mockApp = new MockApplication();  // excelStub と同じインスタンス

const onKeyLog: string[] = [];
(mockApp as any).OnKey = (key: string, procedureName?: string) => {
    onKeyLog.push(`${key}=${procedureName ?? ''}`);
};

let onTimeCallCount = 0;
let lastOnTimeProcName = '';
(mockApp as any).OnTime = (_earliestTime: any, procedureName: string) => {
    onTimeCallCount++;
    lastOnTimeProcName = procedureName;
};

ev.setDefaultBindingObject(mockApp);          // ActiveSheet/Cells はこちら
ev.setBuiltinOverride('Application', mockApp); // Application.XXX はこちら
```

`MockApplication`はnpmパッケージからもexportされている。`VBARunner`の`excelStub`オプションは
`true`の代わりにインスタンスも受け取れるため、npm利用者は型安全なサブクラスとして書ける
(`as any`が不要になる):

```typescript
import { VBARunner, MockApplication } from 'vba-runner';

class AppWithOnKey extends MockApplication {
    onKeyLog: string[] = [];
    OnKey(key: string, procedureName?: string) {
        this.onKeyLog.push(`${key}=${procedureName ?? ''}`);
    }
}

const vbaRunner = new VBARunner('src/vba/KeyHandler.bas', { excelStub: new AppWithOnKey() });
vbaRunner.run('SetupKeyHandlers', []);
// vbaRunner.excelStub は AppWithOnKey のインスタンスそのもの
```

**VBA 定数（`xlUp`・`xlToLeft`・`xlCalculationManual` 等）:**

evaluator に未定義のため、使う定数を手動で注入する:

```typescript
ev.getGlobalEnv().set('xlUp',              -4162);  // Excel の xlUp 定数値
ev.getGlobalEnv().set('xlToLeft',          -4159);
ev.getGlobalEnv().set('xlCalculationManual',   -4135);
ev.getGlobalEnv().set('xlCalculationAutomatic', -4105);
```

VBA 定数の数値は [Microsoft ドキュメント](https://learn.microsoft.com/en-us/office/vba/api/excel.xldirection) を参照。
よく使う値:

| 定数名 | 値 |
|---|---|
| `xlUp` | -4162 |
| `xlDown` | -4121 |
| `xlToLeft` | -4159 |
| `xlToRight` | -4161 |
| `xlCalculationManual` | -4135 |
| `xlCalculationAutomatic` | -4105 |

---

# 📍 Part 1: VBA Runner の MockWorksheet を使う

## 1. クイックスタート

### 最小限のコード例

```typescript
import { MockApplication } from '../../src/engine/mock/MockWorksheet';
import { Lexer } from '../../src/engine/lexer';
import { Parser } from '../../src/engine/parser';
import { Evaluator } from '../../src/engine/evaluator';

// VBA コードの定義
const vbaCode = `
  Function GetTotal() As Long
    GetTotal = Sheets("Data").Range("A1").Value
  End Function
`;

// VBA のコンパイル
const tokens = new Lexer(vbaCode).tokenize();
const ast = new Parser(tokens).parse();
const ev = new Evaluator(() => {});
ev.evaluateModule(ast);

// モック設定
const mockApp = new MockApplication();
mockApp.Sheets('Data').setCellValue('A1', 100);

// Sheets 関数をモックに置き換え
ev.getGlobalEnv().set('Sheets', (name: string) => {
  const sheetObj = mockApp.Sheets(name);
  return {
    Range: (address: string) => sheetObj.Range(address),
  };
});

// テスト実行
const result = ev.callProcedure('GetTotal', []);
console.log(result);  // 100
```

### 実行方法

```bash
npx tsx your-test.ts
```

---

## 2. API リファレンス

### MockApplication

アプリケーションレベルのモック（Excel.Application に相当）。

```typescript
// ワークシートを取得（存在しない場合は自動作成）
Sheets(nameOrIndex: string | number): MockWorksheet

// ワークシート一覧を取得（デバッグ用）
listSheets(): string[]

// すべてをクリア
clear(): void
```

**例**:
```typescript
const mockApp = new MockApplication();
const ws1 = mockApp.Sheets('Sheet1');  // 名前で取得
const ws2 = mockApp.Sheets(1);         // インデックスで取得
```

### MockWorksheet

ワークシートレベルのモック（Worksheet に相当）。

```typescript
// セルに値を設定
// 単一セル: 'A1'
// 範囲: 'A1:C5'
// 配列: [[1,2,3], [4,5,6]]
setCellValue(address: string, value: any): void

// セルから値を取得
getCellValue(address: string): any

// Range オブジェクトを取得
Range(address: string): MockRange

// 全セルをダンプ（デバッグ用）
dump(): Record<string, any>

// ワークシート名を取得
get Name(): string
```

**例**:
```typescript
const ws = mockApp.Sheets('Data');

// 単一セル
ws.setCellValue('A1', 100);
ws.setCellValue('B1', 'Alice');

// 範囲
ws.setCellValue('A1:B3', [
  [1, 'Alice'],
  [2, 'Bob'],
  [3, 'Charlie']
]);

// 読み取り
const val = ws.getCellValue('A1');
const range = ws.Range('A1:B3');
```

### MockRange

セル/範囲を表すオブジェクト（Range に相当）。

```typescript
// セルの値を読み書き
Value: any          // get/set（読み書き両方 Value を経由）
__vbaDefault__: true  // デフォルトプロパティ opt-in マーカー
                      // evaluator が resolveObjectMemberKey(obj, 'value') で Value を解決する
```

**例**:
```typescript
const range = ws.Range('A1');
console.log(range.Value);     // 読み取り
range.Value = 200;             // 書き込み
```

`__vbaDefault__ = true` を持つため、VBA コード内で `x = ws.Range("A1")` と書くと自動的に `.Value` が抽出されます（→ [Q4 参照](#q4-vba-のプロパティメソッドデフォルトメソッドは-javascript-レベルで同じですか)）。

---

## 3. 使用例（5 段階）

### 例 1: 単純な読み書き

```typescript
const mockApp = new MockApplication();
const ws = mockApp.Sheets('Sheet1');

ws.setCellValue('A1', 100);
ws.setCellValue('B1', 200);

const val1 = ws.Range('A1').Value;  // 100
const val2 = ws.Range('B1').Value;  // 200

console.log('[PASS] Simple Read/Write');
```

### 例 2: セル書き込み

```typescript
const mockApp = new MockApplication();
const ws = mockApp.Sheets('TestSheet');

ws.Range('A1').Value = 10;
const original = ws.Range('A1').Value;
console.log(original);  // 10

ws.Range('A1').Value = 50;
const updated = ws.Range('A1').Value;
console.log(updated);  // 50

console.log('[PASS] Cell Write');
```

### 例 3: 範囲操作（配列）

```typescript
const mockApp = new MockApplication();
const ws = mockApp.Sheets('TestSheet');

const data = [[1, 2, 3], [4, 5, 6], [7, 8, 9]];
ws.setCellValue('A1:C3', data);

const range = ws.Range('A1:C3').Value;
console.log(range[0][0]);  // 1
console.log(range[1][1]);  // 5
console.log(range[2][2]);  // 9

console.log('[PASS] Range (Array)');
```

### 例 4: VBA コード + モック統合

```typescript
const vbaCode = `
  Function SumRange() As Long
    Dim data As Variant
    data = Sheets("Data").Range("A1:A5").Value
    Dim sum As Long
    Dim i As Long
    sum = 0
    For i = 1 To UBound(data)
        sum = sum + data(i, 1)
    Next i
    SumRange = sum
  End Function
`;

const tokens = new Lexer(vbaCode).tokenize();
const ast = new Parser(tokens).parse();
const ev = new Evaluator(() => {});
ev.evaluateModule(ast);

const mockApp = new MockApplication();
const ws = mockApp.Sheets('Data');
ws.setCellValue('A1:A5', [[10], [20], [30], [40], [50]]);

ev.getGlobalEnv().set('Sheets', (name: string) => {
  const sheetObj = mockApp.Sheets(name);
  return {
    Range: (address: string) => sheetObj.Range(address),
  };
});

const result = ev.callProcedure('SumRange', []);
console.log(result);  // 150

console.log('[PASS] VBA Code + Mock Integration');
```

### 例 5: 複数シート管理

```typescript
const mockApp = new MockApplication();

mockApp.Sheets('Sheet1').setCellValue('A1', 100);
mockApp.Sheets('Sheet2').setCellValue('A1', 200);
mockApp.Sheets('Sheet3').setCellValue('A1', 300);

const v1 = mockApp.Sheets('Sheet1').Range('A1').Value;  // 100
const v2 = mockApp.Sheets('Sheet2').Range('A1').Value;  // 200
const v3 = mockApp.Sheets('Sheet3').Range('A1').Value;  // 300

console.log('[PASS] Multiple Sheets');
```

---

## 4. テスト実行方法

### tsx で直接実行

```bash
npx tsx tests/spec/your-test.ts
```

### 出力例

```
[Test 1] Simple Read/Write
[PASS] Simple Read/Write

[Test 2] Cell Write
[PASS] Cell Write

[Test 3] Range (Array)
[PASS] Range (Array)

[Test 4] VBA Code + Mock Integration
[PASS] VBA Code + Mock Integration

[Test 5] Multiple Sheets
[PASS] Multiple Sheets

✅ Mock Implementation: 全テスト通過
```

### よくあるエラーと対応

| エラー | 原因 | 対応 |
|--------|------|------|
| `Cannot find module 'Sheets'` | `Sheets` 関数をモックに置き換えてない | `ev.getGlobalEnv().set('Sheets', ...)` を実行 |
| `undefined is not an object` | Range オブジェクトが正しく返されていない | `Range()` が MockRange を返しているか確認 |
| `ReferenceError: Sheets is not defined` | モックが注入されていない | `getGlobalEnv()` を使用しているか確認 |

---

## 5. 対応機能と制限事項

### 対応アドレスフォーマット

| フォーマット | 例 | 説明 |
|-------------|-----|------|
| 単一セル | `A1`, `Z100` | 1 つのセル |
| 範囲 | `A1:C5`, `B2:D10` | 矩形範囲 |
| 絶対参照 | `$A$1`, `$A$1:$B$5` | `$` を除去して通常参照として扱う |
| スピル参照 | `A1#` | `#` を除去して単一セルとして扱う（Excel 365） |
| トリム参照 | `A1.:.A100`, `A1.:A100`, `A1:.A100` | `.` を除去して通常の `:` として扱う（Excel 365） |
| implicit intersection | `@A1`, `@A1:A10` | `@` を除去して通常参照として扱う（Excel 365） |
| Union（カンマ） | `A1:A5,C1:C5` | 複数エリアを縦方向に結合して返す |
| Intersection（スペース） | `A1:C3 B2:D4` | 2 つの範囲の重なり矩形を返す（重なりなしは 0） |
| 列指定 | `A:A` | 対応予定 |
| 行指定 | `1:10` | 対応予定 |

### 現在対応していない機能

| 機能 | 状態 | 説明 |
|------|------|------|
| `Interior.Color` | ❌ | 書式設定（セルの背景色）。getter/setterはあるが値を保持しないno-op |
| `Font` | ❌ | フォント設定 |
| `Offset()` | ❌ | 相対位置参照 |
| `Select()` | ❌ | セル選択 |
| `Copy()` | ❌ | コピー操作 |
| `Paste()` | ❌ | ペースト操作 |
| `UsedRange` | ❌ | 使用範囲 |
| `Application.OnKey` | ❌ | キー入力ハンドラーの登録。呼び出すとエラーになる |
| `Application.OnTime` | ❌ | タイマー実行の登録。呼び出すとエラーになる |

> **`OnKey`/`OnTime` は自前で実装できるが、`Application` を丸ごと差し替える方法は避けること**:
> `MockApplication` をサブクラス化して `OnKey`/`OnTime` を追加し、`excelStub` に `true` の代わりに
> そのインスタンスを渡すのが最も安全（`new VBARunner(path, { excelStub: new MySubclass() })`、
> npmパッケージからも`MockApplication`がexportされているため利用可能）。一時的な上書きで十分なら
> `excelStub` で注入された既存インスタンスに直接メソッドを追加する方法でもよい
> （`mockApp.OnKey = (...) => { ... }`、本節下の [§C](#c-未対応オブジェクトの拡張) 参照）。
> どちらも `ActiveSheet`/`Sheets`/`Range` の解決に使われているのと同一オブジェクトを拡張するだけなので、
> それらの機能を保ったまま `OnKey`/`OnTime` だけ追加できる。
>
> `__mocks__/` 規約(npm パッケージからでも使える、[Part 0](#part-0-mocks-規約で差し替える推奨まずはここを試す)参照)
> の VBA `Class` で `Function Application()` を定義する方法もあるが、これは `Application` 識別子そのものを
> 別オブジェクトに**丸ごと差し替える**ため、対象の VBA コードが同時に `Application.ScreenUpdating` 等の
> 既存サポート済みプロパティへもアクセスしている場合は壊れる。`OnKey`/`OnTime` 以外に `Application` の
> メンバーを一切使わないことが確実な場合のみ許容できる、リスクのある簡易策と理解しておくこと。

### §B: Application プロパティ（ScreenUpdating 等）

#### ⚠️ MockApplication の制限: Application プロパティは「無視」される

`MockApplication` が実装しているのは `Sheets()` メソッドのみ。
`Application.ScreenUpdating`・`Application.Calculation`・`Application.EnableEvents` などのプロパティは**実装されていない**。

しかしエラーにもならない。理由は evaluator の動作にある:

```
VBA:  Application.ScreenUpdating = False
        ↓
evaluator: obj['screenupdating'] = false   ← JS オブジェクトへの動的プロパティ代入
```

JavaScript オブジェクトは未定義プロパティへの代入を無視しないため、エラーなくプロパティが生える。
読み取り時（`state = Application.ScreenUpdating`）も同様に `undefined` が返るだけ。

**結果として何が起きるか:**

| コード | 実際の動作 |
|--------|-----------|
| `screenUpdateState = Application.ScreenUpdating` | `undefined` が代入される |
| `Application.ScreenUpdating = False` | mockApp に `screenupdating: false` が動的に生える（副作用なし） |
| `Application.ScreenUpdating = screenUpdateState` | `undefined` を書き戻す（副作用なし） |

**何がテストされていないか:**
- `ScreenUpdating` の保存・復元ロジックそのものは検証されない
- Excelのパフォーマンス設定（描画停止・手動計算）がテスト中に実際に変わることはない

**運用上の意味:**
スケジューリングやデータ変換のロジックをテストする目的では問題ない。
ただし `Application` を注入しないと、VBA コード内の `Application.ScreenUpdating` 読み取りで
"Type mismatch" が発生し、`AutoScheduleTasks` 自体が実行されない。

### Application プロパティを有効にする拡張例

evaluator は `Application.ScreenUpdating` を `obj['screenupdating']`（小文字）としてアクセスする。
そのため、**小文字のプロパティを持つ拡張クラス**を定義して注入すれば動く。

```typescript
import { MockApplication } from '../../src/engine/mock/MockWorksheet';

// Application.ScreenUpdating / Calculation / EnableEvents に対応した拡張モック
class MockApplicationWithSettings extends MockApplication {
    // evaluator は propName.toLowerCase() でアクセスするため、
    // プロパティ名はすべて小文字で定義する
    screenupdating: boolean = true;
    calculation: any = -4105;   // xlCalculationAutomatic 相当の初期値
    enableevents: boolean = true;
}
```

**テストへの注入:**

```typescript
const mockApp = new MockApplicationWithSettings();
const ws = mockApp.Sheets('Sheet1');

ev.getGlobalEnv().set('ActiveSheet', ws);
ev.getGlobalEnv().set('Application', mockApp);  // ← これを追加

ev.callProcedure('AutoScheduleTasks', []);       // エラーなく実行される

// 保存・復元が正しく行われたことを確認できる
assert.strictEqual(mockApp.screenupdating, true, 'ScreenUpdating が元の値に復元された');
assert.strictEqual(mockApp.enableevents,   true, 'EnableEvents が元の値に復元された');
```

**なぜこれで動くか:**

| VBA コード | evaluator の動作 | 結果 |
|---|---|---|
| `x = Application.ScreenUpdating` | `mockApp['screenupdating']` を読む | `true`（初期値） |
| `Application.ScreenUpdating = False` | `mockApp['screenupdating'] = false` | プロパティが更新される |
| `Application.ScreenUpdating = x` | `mockApp['screenupdating'] = true` | 復元される |

`xlCalculationManual` など VBA 定数は evaluator 内で未定義のため `undefined` になるが、
`calculation` プロパティへの代入・復元のサイクル自体は正常に動作する。

### §C: 未対応オブジェクトの拡張

#### 「こういう機能が欲しい」場合の対応方法

1. **Domain Logic で回避**（推奨）
   - 書式設定に依存したロジックは避ける
   - 色判定は値で判定（例：色コード → 文字列フラグ）

2. **MockWorksheet を拡張**（Part 2 参照）
   - `Interior`, `Font` などのプロパティを追加
   - `Offset()` メソッドを実装

3. **手動テスト**（VBA IDE で実施）
   - 最終確認は実 Excel で検証
   - 自動テストでカバーできない部分

---

# 📍 Part 2: MockWorksheet を参考に自分でモックを実装する

## 6. MockWorksheet.ts の実装を読み解く

### クラス構成

`src/engine/mock/MockWorksheet.ts` の構造：

```
MockApplication（Excel.Application 相当）
  │
  └─ Sheets(nameOrIndex) → Map<string, MockWorksheet>
       │
       └─ MockWorksheet（Worksheet 相当）
            │
            ├─ cells: Map<string, any>（セル値の保存）
            └─ Range(address) → MockRange
                 │
                 └─ MockRange（Range 相当）
                      │
                      └─ Value: any（読み書き可能）
```

### アドレス解析ロジック

MockWorksheet は、アドレス文字列（"A1" や "A1:C5"）を解析して、セル座標に変換：

```typescript
private parseAddress(address: string): { row: number, col: number } | null {
  // "A1" → { row: 1, col: 1 }
  // "Z100" → { row: 100, col: 26 }
}

private parseRange(address: string): { start: {row, col}, end: {row, col} } | null {
  // "A1:C5" → { start: {row:1, col:1}, end: {row:5, col:3} }
}
```

### 2D 配列の扱い

範囲を 2D 配列で設定する場合：

```typescript
ws.setCellValue('A1:B3', [
  [1, 2],
  [3, 4],
  [5, 6]
]);
// 内部的には、各セルに展開されて保存される
// A1=1, B1=2, A2=3, B2=4, ...
```

#### 補足: 単一セルと範囲の違い

`Range()` に渡すアドレスが単一セルか範囲かで、`.Value` の型が変わります：

```typescript
ws.setCellValue('A1', 42);
ws.setCellValue('A1:B3', [[1,2],[3,4],[5,6]]);

ws.Range('A1').Value;      // → 42（スカラー）
ws.Range('A1:B3').Value;   // → [[1,2],[3,4],[5,6]]（2D配列）
```

VBA 側のコードがスカラーを期待しているか 2D 配列を期待しているかは、そのコードが単一セルか範囲を操作しているかで決まります。

#### 補足: 1行・1列の範囲の扱い

**読み取り**時は、1 行または 1 列しかない範囲でも常に 2D 配列を返します（VBA・MockWorksheet ともに同じ）：

```typescript
ws.Range('A1:A4').Value;   // → [[10],[20],[30],[40]]  （4行1列の2D）
ws.Range('A1:D1').Value;   // → [[10,20,30,40]]        （1行4列の2D）
```

**設定**時は、VBA では 1 行/1 列の範囲に 1D 配列を渡せますが、`MockWorksheet` は 2D 配列のみ対応しています：

```vb
' VBA では OK（1D 配列で1列に書き込み）
Range("A1:A4").Value = Array(10, 20, 30, 40)
```

```typescript
// MockWorksheet では 2D で渡す必要がある
ws.setCellValue('A1:A4', [[10],[20],[30],[40]]);
```

#### 補足: 単一値を範囲全体に設定する

スカラー値を渡すと、範囲内の全セルに同じ値が設定されます（VBA・MockWorksheet ともに同じ動作）：

```vb
' VBA
Range("A1:B3").Value = 0   ' 6セルすべてが 0 になる
```

```typescript
// MockWorksheet
ws.setCellValue('A1:B3', 0);  // 6セルすべてが 0 になる
```

#### 補足: 範囲と配列のサイズが異なる場合

| ケース | VBA の挙動 | MockWorksheet の挙動 |
|---|---|---|
| 配列の行数が範囲より少ない | Error 1004（サイズ不一致） | 不足行はスキップ（セル値は変わらない） |
| 配列の列数が範囲より少ない | Error 1004 | 不足列は空文字 `''` で埋める |
| 配列が範囲より大きい | 超過分は無視 | 超過分は無視（同じ） |

テストデータを用意する際は、**範囲と配列のサイズを必ず一致させる**ことを推奨します。サイズ不一致の挙動が VBA と異なるため、意図せず異なる動作になる可能性があります。

#### 補足: 範囲への書き込みは `setCellValue` を使う

VBA では `Range("A1:B3").Value = array` でセルに書き込めますが、MockWorksheet では `ws.Range('A1:B3')` が返すオブジェクトへの `.Value =` はセルに反映されません。その後 `.Value` を読み直しても元の値のままです：

```typescript
ws.setCellValue('A1', 10);
ws.Range('A1:B1').Value = [[99, 99]];         // 書いたつもり
console.log(ws.Range('A1:B1').Value);          // → [[10, 0]]  変わっていない
```

範囲への書き込みは `setCellValue` を使ってください：

```typescript
ws.setCellValue('A1:B1', [[99, 99]]);
console.log(ws.Range('A1:B1').Value);          // → [[99, 99]]  反映される
```

単一セル（`ws.Range('A1').Value = 99`）はセッタが cells マップに接続されているため正常に反映されます。

#### 補足: VBA コードで 2D 配列を扱う場合

VBA 側では 1-based 2 次元配列として受け取ります：

```vb
Dim v As Variant
v = ws.Range("A1:B3").Value   ' v(1,1)〜v(3,2) の2次元配列
Debug.Print v(1, 1)           ' 1行1列目
Debug.Print v(3, 2)           ' 3行2列目
```

テスト側（TypeScript）では 0-based で読み取ります：

```typescript
const v = ws.Range('A1:B3').Value as any[][];
v[0][0];  // 1行1列目
v[2][1];  // 3行2列目
```

### 実装の特徴

1. **セル単位で保存** — 内部は `Map<"A1", value>` で単一セル単位
2. **アドレス正規化** — `"a1"`, `"A1"` どちらでも同じセルを参照
3. **動的 2D 配列変換** — 範囲読み取り時に動的に 2D 配列に変換
4. **複数シート対応** — MockApplication が複数の MockWorksheet を管理

---

## 7. 自作モックのパターン（4 つ） [[→ T-11](REFACTORING_TESTING_CATALOG.md#t-11)]

### パターン 1: JavaScript オブジェクト（最軽量）

**最もシンプル**。テストが単純な場合に最適。

```typescript
const mockSheets = {
  Sales: {
    Range: (address) => ({
      Value: 150000
    })
  }
};

ev.getGlobalEnv().set('Sheets', (name) => mockSheets[name]);
```

VBA コードで `x = Sheets("Sales").Range("A1")` のようにスカラーコンテキストで使う場合は `__vbaDefault__` を追加します：

```typescript
Range: (address) => ({
  __vbaDefault__: true as const,
  Value: 150000,   // getter/setter 両方。x = Range(...) → Value getter、Range(...) = x → Value setter
})
```

**メリット**:
- コードが短い
- 依存関係がない
- セットアップが速い

**デメリット**:
- 複数セルには不向き
- アドレス解析を手動で実装

**使う場合**:
- 単一の値を返すだけ
- テストが 1-2 個のセルだけ

---

### パターン 2: Mock クラス（中程度）

**MockWorksheet を参考に設計**。複数セルの操作が必要な場合。

```typescript
class MockRange {
  readonly __vbaDefault__ = true as const;  // デフォルトプロパティ opt-in

  constructor(private _v: any) {}

  get Value() { return this._v; }   // x = rng → Value getter が呼ばれる
  set Value(v: any) { this._v = v; } // rng = x → Value setter が呼ばれる
}

class MockWorksheet {
  private cells: Map<string, any> = new Map();
  
  Range(address: string) {
    const key = address.toUpperCase();
    if (!this.cells.has(key)) {
      this.cells.set(key, undefined);
    }
    const value = this.cells.get(key);
    return {
      get Value() { return value; },
      set Value(v: any) { this.cells.set(key, v); }
    };
  }
}

class MockApplication {
  private sheets: Map<string, MockWorksheet> = new Map();
  
  Sheets(name: string) {
    if (!this.sheets.has(name)) {
      this.sheets.set(name, new MockWorksheet());
    }
    return this.sheets.get(name)!;
  }
}

// テストで使用
const mockApp = new MockApplication();
ev.getGlobalEnv().set('Sheets', (name) => mockApp.Sheets(name));
```

**メリット**:
- 複数シート対応
- 読み書き可能
- 拡張が容易

**デメリット**:
- コードが増える
- セットアップが複雑

**使う場合**:
- 複数のシート・セルを扱う
- 読み書き両方が必要

---

### パターン 3: ビルダー パターン（初期化が複雑）

初期化データが多い場合、チェーン API で簡潔に書く。

```typescript
class MockSheetBuilder {
  private sheets: Map<string, Map<string, any>> = new Map();
  
  addSheet(name: string): this {
    this.sheets.set(name, new Map());
    return this;
  }
  
  setCellValue(sheetName: string, address: string, value: any): this {
    if (!this.sheets.has(sheetName)) {
      this.sheets.set(sheetName, new Map());
    }
    this.sheets.get(sheetName)!.set(address.toUpperCase(), value);
    return this;
  }
  
  build() {
    return (name: string) => ({
      Range: (address: string) => ({
        Value: this.sheets.get(name)?.get(address.toUpperCase()) ?? 0
      })
    });
  }
}

// テストで使用（初期化がきれい）
const mockSheets = new MockSheetBuilder()
  .addSheet("Sales")
  .setCellValue("Sales", "B2", 150000)
  .setCellValue("Sales", "B3", 200000)
  .setCellValue("Sales", "B4", 175000)
  .build();

ev.getGlobalEnv().set('Sheets', mockSheets);
```

**メリット**:
- 初期化コードが読みやすい
- チェーン API で簡潔
- テストの可読性向上

**デメリット**:
- ビルダークラスの実装が必要
- 一度ビルドするとデータ固定

**使う場合**:
- 複雑な初期化データ
- 複数テストで同じデータを使う

---

### パターン 4: JSON ベース設定（複数テストで共有）

テストデータを JSON で定義して、複数テストで共有。

```typescript
// test-fixtures.json
export const salesDataFixture = {
  Sales: {
    B2: 150000,
    B3: 200000,
    B4: 175000,
    B5: 180000
  },
  Inventory: {
    A1: 100,
    A2: 50,
    A3: 75
  }
};

// test.spec.ts
import { salesDataFixture } from './test-fixtures';

const mockSheets = (sheetName: string) => ({
  Range: (address: string) => ({
    Value: salesDataFixture[sheetName]?.[address.toUpperCase()] ?? 0
  })
});

ev.getGlobalEnv().set('Sheets', mockSheets);
```

**メリット**:
- テストデータを一元管理
- 複数テストで再利用
- テスト実装が簡潔

**デメリット**:
- 別ファイルを管理する必要
- 大規模データは管理困難

**使う場合**:
- 複数テストで同じデータ
- テストシナリオが標準化されている

---

## 8. 実装例：カスタムモックの作成

### 例 1: Interior.Color（セルの背景色）を追加

```typescript
class MockRangeWithFormat {
  constructor(
    private cellValue: any,
    private cellFormat: { color?: number } = {}
  ) {}
  
  get Value() { return this.cellValue; }
  set Value(v: any) { this.cellValue = v; }
  
  get Interior() {
    return {
      Color: this.cellFormat.color,
      set Color(c: number) {
        this.cellFormat.color = c;
      }
    };
  }
}

class MockWorksheetWithFormat {
  private cells: Map<string, MockRangeWithFormat> = new Map();
  
  Range(address: string) {
    const key = address.toUpperCase();
    if (!this.cells.has(key)) {
      this.cells.set(key, new MockRangeWithFormat(undefined));
    }
    return this.cells.get(key)!;
  }
}

// テスト
const ws = new MockWorksheetWithFormat();
ws.Range('A1').Value = 'Important';
ws.Range('A1').Interior.Color = 0xFF0000;  // 赤

console.log(ws.Range('A1').Value);          // 'Important'
console.log(ws.Range('A1').Interior.Color); // 0xFF0000
```

### 例 2: Offset()（相対位置参照）を追加

```typescript
class MockRangeWithOffset {
  constructor(
    private address: string,
    private worksheet: MockWorksheet
  ) {}
  
  get Value() { return this.worksheet.getCellValue(this.address); }
  set Value(v: any) { this.worksheet.setCellValue(this.address, v); }
  
  Offset(rowOffset: number, colOffset: number) {
    // "A1" → { row: 1, col: 1 }
    const [colLetter, rowNum] = this.address.match(/([A-Z]+)(\d+)/)!.slice(1);
    const newRow = parseInt(rowNum) + rowOffset;
    const newCol = String.fromCharCode(colLetter.charCodeAt(0) + colOffset);
    const newAddress = newCol + newRow;
    
    return new MockRangeWithOffset(newAddress, this.worksheet);
  }
}

// テスト
const ws = mockApp.Sheets('Data');
ws.setCellValue('A1', 100);
ws.setCellValue('A2', 200);

const range = ws.Range('A1') as MockRangeWithOffset;
const offsetRange = range.Offset(1, 0);  // A1 + 1行下 = A2
console.log(offsetRange.Value);  // 200
```

### 例 3: 名前付き範囲をサポート

```typescript
class MockWorksheetWithNamedRange {
  private cells: Map<string, any> = new Map();
  private namedRanges: Map<string, string> = new Map();
  
  setNamedRange(name: string, address: string) {
    this.namedRanges.set(name.toUpperCase(), address.toUpperCase());
  }
  
  Range(addressOrName: string) {
    let actualAddress = addressOrName.toUpperCase();
    
    // 名前付き範囲か判定
    if (this.namedRanges.has(actualAddress)) {
      actualAddress = this.namedRanges.get(actualAddress)!;
    }
    
    return {
      get Value() { return this.cells.get(actualAddress); },
      set Value(v: any) { this.cells.set(actualAddress, v); }
    };
  }
}

// テスト
const ws = new MockWorksheetWithNamedRange();
ws.setNamedRange('MyRange', 'A1:B5');
ws.Range('MyRange').Value = 100;
console.log(ws.Range('A1').Value);  // 100
```

---

# 📍 Part 3: モック設計の一般論

## 9. モック設計の原則

### 原則 1: 「完全互換」を目指さない

**❌ 間違い：Excel と完全に同じ挙動を再現**

```typescript
// こんなことはしない：膨大で保守不可能
class CompleteExcelMock {
  // 100+ lines...
  Worksheet: {
    Range: {
      Value: any,
      Interior: { Color: number },
      Font: { Bold: boolean, Size: number, Italic: boolean, Name: string },
      Offset: (r: number, c: number) => Range,
      Copy: () => void,
      Paste: () => void,
      PasteSpecial: (format: number) => void,
      // ... 50+ プロパティ
    }
  }
}
```

**✅ 正解：テストに必要な最小限だけ実装**

```typescript
// 最小限で十分（10-50行で十分）
class LightweightMock {
  Range(address: string) {
    return { Value: 0 };
  }
}
```

**理由**:
- モックが膨大になると、モック自体がバグの源
- 「何のためのテスト？」という本来目的を見失う
- 保守コストが増加

**判定基準**:
- 「テストに必要か？」を問う
- 「Excel との互換性」ではなく「テスト成功」が目的

---

### 原則 2: 動的定義可能にする

**❌ 悪い：事前にすべてのモックを準備**

```typescript
// セットアップが固い
const mockWorksheet = new MockWorksheet();
mockWorksheet.setCellValue('A1', 100);
mockWorksheet.setCellValue('A2', 200);
mockWorksheet.setCellValue('A3', 300);
// ... 100行

// 別のテストではコピー＆ペースト
const mockWorksheet2 = new MockWorksheet();
mockWorksheet2.setCellValue('A1', 500);
// ...
```

**✅ 良い：テスト時に必要なものだけ定義**

```typescript
// 軽量で柔軟
const mockSheets = (name: string) => ({
  Range: (address: string) => ({ Value: 100 })
});

// または
const mockApp = new MockApplication();
// 必要に応じて追加
mockApp.Sheets('Data').setCellValue('A1', 100);
```

**利点**:
- テストごとに異なるデータを簡単に用意
- モック定義がテストの近くにある（可読性向上）
- DRY 原則に従う

---

### 原則 3: 「スタブ」で十分（スパイ機能は不要）

**スタブ** = 値を返すだけ（受動的）  
**スパイ** = 呼び出しを記録（能動的）

**❌ スパイ（初期段階では不要）**

```typescript
// こういった複雑なスパイ機能は不要
const mockMsgBox = jest.fn().mockReturnValue(1);
mockMsgBox('Hello');
expect(mockMsgBox).toHaveBeenCalledWith('Hello');
expect(mockMsgBox).toHaveBeenCalledTimes(1);
```

**✅ スタブ（シンプルで十分）**

```typescript
// テストに必要な値を返すだけ
const mockMsgBox = (message: string) => 1;
const result = mockMsgBox('Hello');
expect(result).toBe(1);
```

**理由**:
- 「呼び出し回数の検証」は、テストが複雑すぎることの兆候
- Domain Logic が分離できていないことを示唆
- テストの目的（戻り値の正確性）に絞る

---

### 原則 4: テストごとにデータを独立させる

複数のテストが1つのモックデータを共有すると、データや期待値の変更が他のテストに波及し、壊れやすいテストスイートになる。**テストごとに入力と期待値を定義する**のが基本方針。

**❌ 悪い：共有モックへの依存**

```typescript
// トップレベルで1つのモックを定義 → 全テストが依存
const mockSheet = new MockWorksheet();
mockSheet.setCellValue('A1', 100);
mockSheet.setCellValue('A2', 200);

test1(mockSheet); // A1=100 を前提
test2(mockSheet); // A2=200 を前提 → test1 の変更が波及する
```

**✅ 良い：テストごとに独立したデータを定義**

```typescript
// テストごとに最小限のデータを用意
it('test1', () => {
    const ws = new MockWorksheet();
    ws.setCellValue('A1', 100);
    // ...
});

it('test2', () => {
    const ws = new MockWorksheet();
    ws.setCellValue('A2', 200);
    // ...
});
```

**理由**:
- どのテストがどのデータに依存するかが明確
- あるテストのデータ変更が他に波及しない
- テスト単独での読解・デバッグが容易

---

## 10. アンチパターン：避けるべき実装

### アンチパターン 1: テスト固有の複雑ロジック

**❌ これはしない：テストコードがビジネスロジック化**

```typescript
// モックが計算を行っている（テストの責務超越）
const mockRange = {
  Value: data.filter(x => x > threshold)
         .map(x => x * 2)
         .reduce((a, b) => a + b, 0)
};
```

**問題**:
- テストのモックが複雑になる
- モックのバグをテストしている状態に
- テスト実装者が計算ロジックを理解する必要

**✅ 正解**:

```typescript
// ビジネスロジックは VBA Function に
// モックはデータ返却だけ
const mockRange = { Value: 500 };  // 計算済みの値を返すだけ

// 計算は VBA で検証
const result = ev.callProcedure('CalculateSum', [100, 200, 200]);
expect(result).toBe(500);
```

---

### アンチパターン 2: グローバルモック状態

**❌ これはしない：テスト間で状態が共有**

```typescript
// グローバル状態：テスト分離が破れる
let globalMockState = { value: 0 };

it('test 1', () => {
  globalMockState.value = 100;
  // ...
});

it('test 2', () => {
  // test 1 の副作用に依存（テスト順序に依存）
  expect(globalMockState.value).toBe(100);
});
```

**問題**:
- テスト実行順序に依存
- テスト分離の原則を破る
- デバッグが困難

**✅ 正解**:

```typescript
it('test 1', () => {
  const mockApp = new MockApplication();  // 新しいインスタンス
  mockApp.Sheets('Data').setCellValue('A1', 100);
  // ...
});

it('test 2', () => {
  const mockApp = new MockApplication();  // 独立したインスタンス
  mockApp.Sheets('Data').setCellValue('A1', 200);
  // ...
});
```

---

### アンチパターン 3: モック実装に時間を使い過ぎ

**❌ これはしない：モック実装が本体より長い**

```typescript
// VBA コード: 50行
// モック実装: 500行  ← 危険信号！
```

**兆候**:
- モック実装が VBA コードより長い
- 複数の特殊ケースを処理している
-「テストが複雑すぎて理解できない」

**✅ 正解**:

```
モック実装 < VBA コード  が原則

長い場合は：
1. ビジネスロジックを分離（REFACTORING_GUIDE.md）
2. よりシンプルなモックパターンを採用
```

---

## 11. 判断フロー：いつモック化すべきか

### チェックリスト

モック実装が必要か判定：

- [ ] **Domain Logic を Function に分離できたか？**
  - 「はい」→ 以下に進む
  - 「いいえ」→ REFACTORING_GUIDE.md を参照してリファクタリング

- [ ] **I/O は Sub に集約できたか？**
  - 「はい」→ 以下に進む
  - 「いいえ」→ REFACTORING_GUIDE.md を参照

- [ ] **テスト対象の Sub は 20 行以内か？**
  - 「はい」→ 以下に進む
  - 「いいえ」→ さらに分割を検討

- [ ] **Excel オブジェクトの使用箇所は 5 個以内か？**
  - 「はい」→ モック実装を開始
  - 「いいえ」→ さらなるリファクタリング

### 判定結果

| 結果 | 対応 |
|------|------|
| **すべて Yes** | モック実装開始（このガイド Part 1） |
| **No が 1-2 個** | リファクタリング（REFACTORING_GUIDE.md） |
| **No が 3 個以上** | 大幅なリファクタリングが必要 |

---

## 12. テスト戦略：段階的なテスト化 [[→ S-01](REFACTORING_TESTING_CATALOG.md#s-01)]

```
┌─────────────────────────────────────┐
│ Unit Tests（単体テスト）             │
├─────────────────────────────────────┤
│ Domain Logic Functions のみ           │
│ → Excel 依存なし                     │
│ → 100% カバー可能                    │
│ テスト数：多い（関数ごと）            │
└─────────────────────────────────────┘
           ↓
┌─────────────────────────────────────┐
│ Integration Tests（統合テスト）      │
├─────────────────────────────────────┤
│ 複数関数 + I/O Sub                   │
│ → Mock Worksheet を使用              │
│ → Excel なしで実行                   │
│ テスト数：中程度（シナリオごと）     │
└─────────────────────────────────────┘
           ↓
┌─────────────────────────────────────┐
│ E2E Tests（エンドツーエンド）        │
├─────────────────────────────────────┤
│ 実際の Excel ファイルで実行           │
│ → VBA IDE で手動実行                 │
│ → UI 操作の最終確認                  │
│ テスト数：少ない（重要ケースのみ）   │
└─────────────────────────────────────┘
```

**推奨される配分**:
- Unit Tests: 70-80%（大部分）
- Integration Tests: 15-25%（主要シナリオ）
- E2E Tests: 5-10%（最終確認）

---

## 13. よくある質問（FAQ）

### Q1: MockWorksheet に機能を追加したい

**A**: Part 2 の「実装例：カスタムモックの作成」を参照してください。

例えば、`Interior.Color` を追加したい場合：

```typescript
class MockRangeWithFormat {
  constructor(private value: any) {}
  
  get Interior() {
    return { Color: 0xFF0000 };
  }
}
```

拡張ポイント：
- `Interior`, `Font` などのプロパティ追加
- `Offset()`, `Select()` などのメソッド追加
- `UsedRange` などの特殊な Range を実装

---

### Q2: 複数のテストで同じモック状態を使いたい

**A**: パターン 4（JSON ベース設定）を使用してください。

```typescript
// test-fixtures.ts
export const commonTestData = {
  Sales: { A1: 100, A2: 200 },
  Inventory: { A1: 50, A2: 75 }
};

// test1.spec.ts
import { commonTestData } from './test-fixtures';
const mockSheets = (name) => ({
  Range: (addr) => ({ Value: commonTestData[name][addr] })
});

// test2.spec.ts も同じ mockSheets を使用
```

---

### Q3: VBA コードのデバッグが難しい

**A**: 以下の方法を組み合わせてください：

1. **小さな Function ごとにテスト**
   ```typescript
   const result = ev.callProcedure('CalcLevel', [1, 3, 0.3]);
   console.log(result);
   ```

2. **モックデータを print して確認**
   ```typescript
   console.log('[DEBUG] mockApp state:', mockApp.dump());
   ```

3. **最終的には実 Excel で確認**
   - VBA IDE でステップ実行
   - Excel シートに手動でデータ入力

---

### Q4: VBA のプロパティ・メソッド・デフォルトメソッドは JavaScript レベルで同じですか？

**A**: VBA Runner の実装上、**すべて JavaScript のプロパティアクセスまたは関数呼び出しに統一されます**。

| VBA の構文 | JavaScript での実装 |
|---|---|
| `rng.Value`（プロパティ読み取り） | `rng.value`（小文字化したプロパティ） |
| `rng.Value = 1`（プロパティ書き込み） | `rng.value = 1` |
| `rng.ClearContents()`（メソッド呼び出し） | `rng.clearcontents()` |
| `v = rng`（スカラーコンテキスト → デフォルトメンバー） | `rng.value` へのアクセスに変換 |
| `Set v = rng`（オブジェクトコンテキスト） | `rng` 自身（オブジェクト参照） |

VBA では「プロパティ」「メソッド」「デフォルトメンバー」の区別がありますが、JavaScript オブジェクトにはその区別がありません。モックを実装する際は次の点だけ押さえれば十分です：

- **読み取り専用プロパティ** → getter か単純な値で定義
- **書き込みが必要なプロパティ** → 通常のフィールドとして定義
- **デフォルトメンバー** → スカラーコンテキスト（`v = rng`）では VBA Runner が `Value` getter を呼んで値を抽出する。ただし opt-in が必要（下記参照）。`Set v = rng` のようなオブジェクトコンテキストではオブジェクト自体が代入されるため変換は起きない

#### スカラーコンテキストの opt-in: `__vbaDefault__`

`v = rng`（Let 代入）で `.Value` が自動抽出されるには、モックオブジェクトに以下を実装します：

```typescript
class MockRange {
    readonly __vbaDefault__ = true as const;  // opt-in マーカー

    constructor(private _v: any) {}

    get Value(): any { return this._v; }   // v = rng → Value getter が呼ばれる
    set Value(val: any) { this._v = val; } // rng = x → Value setter が呼ばれる
}
```

evaluator は読み書き両方で `resolveObjectMemberKey(obj, 'value')` を使って `Value` プロパティを解決します。
`__vbaDefault__` を持たないオブジェクトは VBA Runner がスカラー抽出を行わないため、`v = obj` でもオブジェクトがそのまま代入されます。`MockRange` はすでにこの実装を持っています。

---

### Q5: 書式設定（Color, Font）が必要な場合は？

**A**: 3 つの対応方法：

1. **Domain Logic で回避（推奨）**
   - 色判定の代わりに値で判定
   - 例：赤 → 負の値、緑 → 正の値

2. **MockWorksheet を拡張（Part 2 参照）**
   - `Interior.Color` を実装
   - テスト可能な形に

3. **手動テスト（VBA IDE）**
   - 自動テストでカバーできない部分
   - 実 Excel で最終確認

---

### Q5: Union Range（複数の離れた範囲）を使いたい

**A**: MockWorksheet は Union（カンマ区切り）と Intersection（スペース区切り）の両方に対応しています。

**Union: カンマ区切り**
```typescript
// setCellValue で複数エリアに一括設定
ws.setCellValue('A1:A2,C1:C2', 99);

// Range().Value で複数エリアを縦に結合して取得
const v = ws.Range('A1:A2,C1:C2').Value as any[][];
// v[0][0] = A1, v[1][0] = A2, v[2][0] = C1, v[3][0] = C2
```

**Intersection: スペース区切り**
```typescript
ws.setCellValue('A1:C3', 1);
ws.setCellValue('B2', 99);
const v = ws.Range('A1:C3 B2:D4').Value as any[][];
// A1:C3 と B2:D4 の重なりは B2:C3
// 重なりがない場合は 0 を返す
```

---

---

## 14. Declare ステートメントと DLL 呼び出しのスタブ差し替え

### Declare の自動スタブ化

`Private Declare Function` / `Private Declare Sub` は DLL や共有ライブラリを呼び出す宣言です。
VBA Runner は宣言を評価した時点で **自動的にスタブ関数として登録** します。

```vb
Private Declare Function GetTickCount Lib "kernel32" () As Long
```

登録されるスタブの挙動:
- 実際の DLL は呼ばれない
- 常に `0` を返す
- `[DECLARE STUB] Calling Function GetTickCount from "kernel32" (Alias: N/A)` をコンソールに出力する

`Declare` を含むソースをそのまま読み込んでもエラーにはならず実行できます。
ただし戻り値が常に `0` なので、その関数の戻り値に依存するロジックは正しく動きません。

### VBARunner からスタブを登録する

`VBARunner` が持つ `set(name, value)` を使うと、ソースを改変せずに DLL 関数を差し替えられます。
コンストラクターでソースを読み込んだ後に呼ぶだけで、複数ソースをまたいでグローバルに有効になります。

```typescript
const vbaRunner = new VBARunner('MyLib.bas');

// コンストラクタ（ソース読み込み）の後に登録する
// Private Declare でもグローバル env に登録されるため差し替え可能
vbaRunner.set('gettickcount', () => 12345);
```

> **注意**: `set()` はコンストラクタの後に呼んでください。コンストラクタ内でソースが評価されて `Declare` が自動スタブとして登録されるため、その前に `set()` しても上書きされます。

識別子はすべて小文字に正規化されるため、`GetTickCount` → `'gettickcount'` として登録します。

### 具体例：DLL 関数に依存するロジックのテスト

```vb
' MyLib.bas（テスト対象）
Private Declare Function GetElapsedMs Lib "myutil.dll" () As Long

Function IsTimeout(limitMs As Long) As Boolean
    IsTimeout = (GetElapsedMs() > limitMs)
End Function
```

```typescript
const vbaRunner = new VBARunner('MyLib.bas');

// DLL 関数を TypeScript のスタブで差し替える（GetElapsedMs → 小文字で登録）
vbaRunner.set('getelapsedms', () => 5000);  // 5000ms 経過したと見なす

// DLL なしで IsTimeout のロジックをテストできる
assert.strictEqual(vbaRunner.run('IsTimeout', [3000]), true,  '5000 > 3000 → timeout');
assert.strictEqual(vbaRunner.run('IsTimeout', [9000]), false, '5000 < 9000 → not timeout');
```

### CreateObject 系との使い分け

| 依存の種類 | 対応方法 |
|---|---|
| `Declare Function` / DLL 直呼び出し | 同名 VBA 関数を後ろに追記して差し替え（本節） |
| `CreateObject("Foo.Bar")` | `vbaRunner.registerComObject(factory)` で TypeScript ファクトリを登録（REFERENCE.md 参照） |

---

## 15. 関連ドキュメント

- **`docs/TESTING_STRATEGY.md`** — テスト設計の原則（最初に読む）
- **`docs/REFACTORING_GUIDE.md`** — Domain Logic 分離（モック前に読む）
- **`docs/INTEGRATION_TEST_EXAMPLE.md`** — 統合テストの実装例
- **`docs/TEST_FRAMEWORK_GUIDE.md`** — JavaScript テストフレームワーク活用
- **`tests/spec/mock-usage.test.ts`** — VBA Runner のテスト実装例
- **`src/engine/mock/MockWorksheet.ts`** — MockWorksheet の実装ソース

---

## まとめ

### モック実装の原則

```
Domain Logic 分離（最優先）
    ↓ 分離不可能な場合のみ
軽量モック実装（このガイド）
    ↓ さらに複雑な場合
Excel 統合テスト（VBA IDE での手動テスト）
```

### 重要な心がけ

1. **モックは最後の手段** — 可能な限り分離を試みる
2. **軽量を心がける** — 50 行のモックで十分
3. **テスト責務を混同しない** — モックは値を返すだけ
4. **テスト分離を守る** — 各テストは独立したモック

これらの原則に従えば、効果的で保守性の高いテストが実現できます。
