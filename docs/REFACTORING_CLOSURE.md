# クロージャによるループ抽象化 — VBA での実現

VBA にはクロージャ（関数と状態のセット）がなく、高階関数も持てない。
しかしクロージャが解決する問題は **クラス** で同等に表現できる（Java のインナークラスと同じアプローチ）。

ここでは「キーブレイクによるグループ別小計レポート」を例に、JavaScript のクロージャ版から VBA のクラス版への読み替えを示す。JavaScript によるシンプルな実装と、VBA によるソース量は多いが実務で役に立つ実例の両面を示すことが目的である。

---

## 例: 部門別売上小計レポート

次のデータを部門（`dept`）ごとにグループ化し、各グループの合計金額を出力する。
データはキー（`dept`）でソート済みであることが前提。

```
営業  100
営業  200
開発  150
開発  300
総務   80
```

期待する出力:
```
営業: 300
開発: 450
総務: 80
```

---

## JavaScript: ハードコード版

ループ構造と処理の中身が一体になっている。別のキーや別の集計ロジックに使い回せない。

```javascript
function reportByDept(records) {
    let currentDept = undefined;
    let subtotal = 0;

    for (const r of records) {
        if (r.dept !== currentDept) {
            if (currentDept !== undefined) {
                console.log(`${currentDept}: ${subtotal}`);  // ブレーク時の処理
            }
            currentDept = r.dept;
            subtotal = 0;
        }
        subtotal += r.amount;  // 各レコードの処理
    }
    if (currentDept !== undefined) {
        console.log(`${currentDept}: ${subtotal}`);  // 最終グループ
    }
}
```

---

## JavaScript: クロージャ版

ループ構造（「キーが変わったら onBreak を呼ぶ」）を `runKeyBreak` 関数に切り出し、
処理の中身をクロージャで渡す。`subtotal` はクロージャが捕捉したスコープ変数として共有される。

```javascript
// ループ構造だけを担う汎用関数
function runKeyBreak(records, getKey, onRecord, onBreak) {
    let currentKey = undefined;
    for (const r of records) {
        const key = getKey(r);
        if (key !== currentKey) {
            if (currentKey !== undefined) onBreak(currentKey);
            currentKey = key;
        }
        onRecord(r);
    }
    if (currentKey !== undefined) onBreak(currentKey);
}

// 呼び出し側: subtotal はクロージャが閉じ込めた状態
function reportByDept(records) {
    let subtotal = 0;
    runKeyBreak(
        records,                          // records
        r => r.dept,                      // getKey — キー抽出
        r => { subtotal += r.amount; },   // onRecord — 各レコードの処理
        key => {                          // onBreak
            console.log(`${key}: ${subtotal}`);
            subtotal = 0;                 // 次グループへリセット
        }
    );
}
```

`runKeyBreak` は汎用なので、別のキーや別の集計ロジックにそのまま使い回せる。

---

| | ハードコード版 | クロージャ版 |
|---|---|---|
| ループ構造の再利用 | ❌ | ✅ |
| ストリーム処理 | ✅ | ✅ |
| 可読性 | △ | ◎ |
| VBA への翻訳しやすさ | △ | ✅ |

---

## VBA での実現: 2 つのアプローチ

VBA でクロージャ相当を実現する方法は主に 2 つある。

| | **方法1: CallByName（duck-typing）** | **方法2: Implements（interface）** |
|---|---|---|
| インターフェース定義 | 不要 | 必要（`.cls` ファイル） |
| 型チェック | なし（実行時エラー） | コンパイル時チェック |
| メソッド呼び出し | `CallByName obj, "Method", VbMethod` | `obj.Method` |
| ファイル数 | 少ない | やや多い |
| 向いている場面 | 小規模・試作・既存クラスの流用 | 中〜大規模・チーム開発 |

どちらの方法でも `RunKeyBreak` の呼び出し側は同じ形になる:

```vba
Dim body As New DeptSubtotalBody
RunKeyBreak records, body
```

---

## 方法1: CallByName による duck-typing（`closure1/`）

インターフェース定義クラスを用意せず、メソッド名の文字列で呼び出す。
`body` は `As Object` 型で受け取るため、`GetKey` / `OnRecord` / `OnBreak` を持つ任意のクラスをそのまま渡せる。

### ループ構造（`KeyBreakModule.bas`）

```vba
Public Sub RunKeyBreak(records As Collection, body As Object)
    Dim currentKey As Variant
    Dim isFirst As Boolean
    isFirst = True

    Dim rec As Object
    For Each rec In records
        Dim key As Variant
        key = CallByName(body, "GetKey", VbMethod, rec)

        If isFirst Or key <> currentKey Then
            If Not isFirst Then CallByName body, "OnBreak", VbMethod, currentKey
            currentKey = key
            isFirst = False
        End If
        CallByName body, "OnRecord", VbMethod, rec
    Next rec

    If Not isFirst Then CallByName body, "OnBreak", VbMethod, currentKey
End Sub
```

### 処理クラス（`DeptSubtotalBody.cls`）

```vba
Option Explicit

Private mSubtotal As Long
Private mResult As String

Public Function GetKey(rec As Object) As Variant
    GetKey = rec.Department
End Function

Public Sub OnRecord(rec As Object)
    mSubtotal = mSubtotal + rec.Amount
End Sub

Public Sub OnBreak(currentKey As Variant)
    mResult = mResult & currentKey & ": " & mSubtotal & vbLf
    mSubtotal = 0
End Sub

Public Function GetResult() As String
    GetResult = mResult
End Function
```

オブジェクトが `GetKey` / `OnRecord` / `OnBreak` の3つのメソッドを持っていさえすれば `RunKeyBreak` に渡せる。

---

## 方法2: Implements によるインターフェース（`closure2/`）

インターフェースクラス `IKeyBreakBody` を定義し、実装クラスに `Implements IKeyBreakBody` を宣言する。
`body` は `As IKeyBreakBody` 型で受け取るため、コンパイル時に型チェックが効く。

### Step 1: インターフェースを定義する（`IKeyBreakBody.cls`）

```vba
Option Explicit

Public Function GetKey(rec As Object) As Variant
End Function

Public Sub OnRecord(rec As Object)
End Sub

Public Sub OnBreak(currentKey As Variant)
End Sub
```

### Step 2: ループ構造を汎用プロシージャとして書く（`KeyBreakModule.bas`）


```vba
' body の型としてインターフェース IKeyBreakBody を指定できる
' インターフェースによって3つのメソッドが定義されていることを保証する
Public Sub RunKeyBreak(records As Collection, body As IKeyBreakBody)
    Dim currentKey As Variant
    Dim isFirst As Boolean
    isFirst = True

    Dim rec As Object
    For Each rec In records
        Dim key As Variant
        key = body.GetKey(rec)

        If isFirst Or key <> currentKey Then
            If Not isFirst Then body.OnBreak currentKey
            currentKey = key
            isFirst = False
        End If
        body.OnRecord rec
    Next rec

    If Not isFirst Then body.OnBreak currentKey
End Sub
```

### Step 3: 処理クラスを実装する（`DeptSubtotalBody.cls`）

```vba
Option Explicit
Implements IKeyBreakBody

Private mSubtotal As Long
Private mResult As String

Private Function IKeyBreakBody_GetKey(rec As Object) As Variant
    IKeyBreakBody_GetKey = rec.Department
End Function

Private Sub IKeyBreakBody_OnRecord(rec As Object)
    mSubtotal = mSubtotal + rec.Amount
End Sub

Private Sub IKeyBreakBody_OnBreak(currentKey As Variant)
    mResult = mResult & currentKey & ": " & mSubtotal & vbLf
    mSubtotal = 0
End Sub

Public Function GetResult() As String
    GetResult = mResult
End Function
```

### Step 4: 呼び出し側（`ReportModule.bas`）

```vba
Sub PrintDeptReport(records As Collection)
    Dim body As New DeptSubtotalBody
    RunKeyBreak records, body
End Sub
```

---

## JavaScript クロージャ版との対応

| JavaScript | 方法1 (CallByName) | 方法2 (Implements) |
|---|---|---|
| `r => r.dept` | `GetKey` メソッド（直接呼び出し） | `IKeyBreakBody_GetKey` メソッド |
| `r => { subtotal += r.amount }` | `OnRecord` メソッド（直接呼び出し） | `IKeyBreakBody_OnRecord` メソッド |
| `key => { print; subtotal = 0 }` | `OnBreak` メソッド（直接呼び出し） | `IKeyBreakBody_OnBreak` メソッド |
| `let subtotal = 0`（捕捉変数） | `Private mSubtotal As Long` | `Private mSubtotal As Long` |
| 型チェック | なし | コンパイル時 |

---

## 別のグループ処理への使い回し

`RunKeyBreak` はそのままで、body クラスだけ差し替えれば別の集計に対応できる（どちらの方法も同じ）。

```vba
' 例: 部門ごとのレコード件数をカウントする
Private mCount As Long

Public Function GetKey(rec As Object) As Variant  ' 方法1
    GetKey = rec.Department
End Function

Public Sub OnRecord(rec As Object)
    mCount = mCount + 1
End Sub

Public Sub OnBreak(currentKey As Variant)
    mResult = mResult & currentKey & ": " & mCount & "件" & vbLf
    mCount = 0
End Sub
```

```vba
' 呼び出し側を変えるだけ
Sub PrintDeptCount(records As Collection)
    Dim body As New DeptCountBody
    RunKeyBreak records, body
End Sub
```

---

## クラスで実現する上での注意点

- **インターフェースメソッドは必ずすべて実装する**: 1つでも欠けるとコンパイルエラーになる（方法2の強み）。
- **インターフェース型変数から独自メソッドにはアクセスできない**: `Dim body As IKeyBreakBody` で宣言した変数から `GetResult` など独自メソッドを呼ぶにはコンクリート型の変数を別途用意する必要がある（方法2の制約）。
- **CallByName はリフレクション**: メソッド名をタイプミスしても実行時までエラーにならない（方法1の弱点）。
- **前処理・後処理の共通化**: `OnRecord` / `OnBreak` に加え `OnStart` (最初のレコード前) を加えると初期化処理もカプセル化できる。

---

## まとめ: VBA における重複コードへの処方箋

VBA にはクロージャがないため、「ループの構造は同じだが中身の処理だけ違う」コードを書くとき、
開発者はしばしば似たようなループを何度もコピーして微妙に書き変えるという対応をとりがちである。
これはループに限らず、「前処理と後処理は共通で、途中の計算や出力だけが異なる」パターン全般に当てはまる。

こうした重複に対して、本稿で示した **クラスによるメソッド委譲** は有効な手法となる。
方法1（CallByName）は手軽でファイル数も少なく、既存クラスをそのまま流用できる。
方法2（Implements）はコンパイル時の型チェックが効き、IDE の補完も受けられるため、チーム開発や規模が大きくなった場合に安心感がある。

ループ処理以外にも、たとえば次のような場面で同じアプローチが応用できる:

- **レポート出力**: 出力先（Debug.Print / ファイル / シート書き込み）だけ差し替えたい
- **バリデーション**: チェック項目のリストは共通で、各チェックのロジックだけ違う
- **変換処理**: データの読み取り・書き込みは共通で、変換ルールだけ差し替えたい

---

## サンプルソース

### 方法1: CallByName（`sample/src/vba/closure1/`）

| ファイル | 役割 |
|---------|------|
| `KeyBreakModule.bas` | 汎用ループ構造 `RunKeyBreak`（CallByName 版） |
| `DeptSubtotalBody.cls` | 売上小計クラス（GetKey / OnRecord / OnBreak） |
| `DeptCountBody.cls` | 件数カウントクラス |
| `SalesRecord.cls` | テスト用データクラス |
| `ClosureTest.bas` | テストドライバー |

```bash
./vba-run.sh sample/src/vba/closure1/ClosureTest.bas RunSubtotalTest
./vba-run.sh sample/src/vba/closure1/ClosureTest.bas RunCountTest
```

### 方法2: Implements（`sample/src/vba/closure2/`）

| ファイル | 役割 |
|---------|------|
| `IKeyBreakBody.cls` | インターフェース（GetKey / OnRecord / OnBreak） |
| `KeyBreakModule.bas` | 汎用ループ構造 `RunKeyBreak`（Implements 版） |
| `DeptSubtotalBody.cls` | 売上小計クラス（IKeyBreakBody を実装） |
| `DeptCountBody.cls` | 件数カウントクラス（IKeyBreakBody を実装） |
| `SalesRecord.cls` | テスト用データクラス |
| `ReportModule.bas` | 呼び出し側 |
| `ClosureTest.bas` | テストドライバー |

```bash
./vba-run.sh sample/src/vba/closure2/ClosureTest.bas RunSubtotalTest
./vba-run.sh sample/src/vba/closure2/ClosureTest.bas RunCountTest
```

---

## 参照

- [REFACTORING_TESTING_CATALOG.md — R-10](REFACTORING_TESTING_CATALOG.md#r-10)
