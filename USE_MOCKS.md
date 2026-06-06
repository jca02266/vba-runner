# __mocks__ 使用ガイド

VBA コードが `MsgBox` や `Range` といった Excel / Windows 環境に依存する処理を呼び出すとき、
エンジン単体での実行はエラーになります。`__mocks__/` ディレクトリ（または `__mocks__.*` ファイル）に
モックを置くと、`VBARunner` / 拡張機能の `▶ Run` 実行時にエンジンのビルトインが自動的に差し替えられ、
Excel がなくても VBA コードを実行できるようになります。

## モック形式の選び方

モックは `.bas`（VBA）・`.js`（JavaScript）・`.ts`（TypeScript）の 3 形式で書けます。
どれを選ぶかは「何をモックしたいか」と「どの環境で動かすか」で決まります。

| 形式     | 追加インストール              | JS ライブラリ利用               | 向いている用途                           |
| ------ | --------------------- | ------------------------ | --------------------------------- |
| `.bas` | 不要                    | **不可**（VBA の範囲内のみ）       | Excel オブジェクトの構造スタブ・単純なロジック差し替え    |
| `.js`  | 不要                    | 可（Node.js API・npm パッケージ） | COM 依存の代替・正規表現・HTTP など JS ネイティブ機能 |
| `.ts`  | tsx 等が必要（VS Code は任意） | 可（型安全）                   | `.js` と同じことを型チェック付きで書きたい場合        |

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

## 1. MsgBox をコンソール出力に差し替える

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

## 2. VBA クラスで Excel オブジェクトをスタブする

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
（→ [3. ビルトインをラップする](#3-ビルトインをラップする構想)）。

***

## 3. ビルトインをラップする（構想）

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

## 4. 正規表現ライブラリのモック（JS モックが必要な例）

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

## 用語について

「モック」という言葉は文脈によって意味が異なります。

狭義の「モック」はテストダブルの一種であり、「どのメソッドが何回呼ばれたか」「どんな引数で呼ばれたか」を事前に期待値として設定し、事後に検証できるオブジェクトを指します。この定義に照らすと、本ドキュメントの例は厳密には「スタブ」（決まった値を返すだけの最小実装）や「フェイク」（本物と同様に動く代替実装）であり、狭義の「モック」の説明としては不適切です。

一方、Jest の `__mocks__/` ディレクトリの慣習に代表されるように、テストダブル全般を広く「モック」と呼ぶ用法も広く普及しています。本ドキュメントはこの広義の「モック」の意味でこの言葉を使っています。
