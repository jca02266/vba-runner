# クロージャによるループ抽象化 — VBA での実現

VBA にはクロージャ（関数と状態のセット）がなく、高階関数も持てない。
しかしクロージャが解決する問題は **インターフェースとクラス** で同等に表現できる（Java のインナークラスと同じアプローチ）。

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

ループ構造（「キーが変わったら onBreak を呼ぶ」）を `keyBreak` 関数に切り出し、
処理の中身をクロージャで渡す。`subtotal` はクロージャが捕捉したスコープ変数として共有される。

```javascript
// ループ構造だけを担う汎用関数
function keyBreak(records, getKey, onRecord, onBreak) {
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
    keyBreak(
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

`keyBreak` は汎用なので、別のキーや別の集計ロジックにそのまま使い回せる。

---

## JavaScript: groupBy 版

ES2024 で導入された `Object.groupBy()` は**組み込み高階関数**であり、ループ抽象が言語レベルで提供されているため簡潔に書ける。
ただしデータをすべてメモリに展開するため、大量データには向かない。
また「ソート済みデータをストリーム的に処理する」という前提が失われる点に注意。

```javascript
function reportByDept(records) {
    const groups = Object.groupBy(records, r => r.dept);
    for (const [dept, items] of Object.entries(groups)) {
        const subtotal = items.reduce((sum, r) => sum + r.amount, 0);
        console.log(`${dept}: ${subtotal}`);
    }
}
```

| | ハードコード版 | クロージャ版 | groupBy 版 |
|---|---|---|---|
| ループ構造の再利用 | ❌ | ✅ | △（reduce として） |
| ストリーム処理 | ✅ | ✅ | ❌（全件メモリ展開） |
| 可読性 | △ | ◎ | ◎ |
| VBA への翻訳しやすさ | △ | ✅ | △ |

---

## VBA: クラスを使ったクロージャ代替

VBA ではクロージャの代わりに **インターフェース + クラス** を使う（Java の匿名クラスと同じ構造）。
「処理の中身」をインターフェースとして定義し、呼び出しごとに実装クラスを差し替える。

### Step 1: インターフェースを定義する（`IKeyBreakBody.cls`）

```vba
Option Explicit

' 各レコードを処理する
Public Sub OnRecord(rec As Object)
End Sub

' キーブレイク発生時（グループ終了時）の処理
Public Sub OnBreak(currentKey As Variant)
End Sub
```

### Step 2: ループ構造を汎用プロシージャとして書く（`KeyBreakModule.bas`）

```vba
Option Explicit

' ループ構造だけを担う汎用プロシージャ
' records : Object の Collection（キーでソート済み）
' keyProp : キーとして使うプロパティ名（CallByName で取得）
' body    : IKeyBreakBody の実装クラス
Public Sub RunKeyBreak(records As Collection, keyProp As String, body As IKeyBreakBody)
    Dim currentKey As Variant
    Dim isFirst As Boolean
    isFirst = True

    Dim rec As Object
    For Each rec In records
        Dim key As Variant
        key = CallByName(rec, keyProp, VbGet)

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

### Step 3: 処理の中身をクラスで実装する（`DeptSubtotalBody.cls`）

```vba
Option Explicit
Implements IKeyBreakBody

Private mSubtotal As Long

' 各レコード: 金額を累積
Private Sub IKeyBreakBody_OnRecord(rec As Object)
    mSubtotal = mSubtotal + rec.Amount
End Sub

' グループ終了: 小計を出力してリセット
Private Sub IKeyBreakBody_OnBreak(currentKey As Variant)
    Debug.Print currentKey & ": " & mSubtotal
    mSubtotal = 0
End Sub
```

### Step 4: 呼び出し側（`ReportModule.bas`）

```vba
Sub PrintDeptReport(records As Collection)
    Dim body As New DeptSubtotalBody
    RunKeyBreak records, "Department", body
End Sub
```

JavaScript のクロージャ版と対応させると:

| JavaScript | VBA |
|---|---|
| `keyBreak(records, getKey, onRecord, onBreak)` | `RunKeyBreak records, "Department", body` |
| `r => r.dept`（キー抽出クロージャ） | `keyProp = "Department"`（プロパティ名文字列 + `CallByName`） |
| `r => { subtotal += r.amount }`（クロージャ） | `IKeyBreakBody_OnRecord`（クラスメソッド） |
| `key => { print; subtotal = 0 }`（クロージャ） | `IKeyBreakBody_OnBreak`（クラスメソッド） |
| `let subtotal = 0`（捕捉された変数） | `Private mSubtotal As Long`（クラスのフィールド） |

---

## 別のグループ処理への使い回し

`RunKeyBreak` はそのままで、`IKeyBreakBody` の実装クラスだけ差し替えれば別の集計に対応できる。

```vba
' 例: 部門ごとのレコード件数をカウントする実装
Option Explicit
Implements IKeyBreakBody

Private mCount As Long

Private Sub IKeyBreakBody_OnRecord(rec As Object)
    mCount = mCount + 1
End Sub

Private Sub IKeyBreakBody_OnBreak(currentKey As Variant)
    mResult = mResult & currentKey & ": " & mCount & "件" & vbLf
    mCount = 0
End Sub

Public Function GetResult() As String
    GetResult = mResult
End Function
```

```vba
' 呼び出し側を変えるだけ
Sub PrintDeptCount(records As Collection)
    Dim body As New DeptCountBody
    RunKeyBreak records, "Department", body
End Sub
```

---

## クラスで実現する上での注意点

- **Implements はコンパイル時チェックのみ**: インターフェースを `Implements` しても、実行時に正しく動くかどうかはテストが必要。VBA IDE は型安全性を保証しない場面がある
- **インターフェースメソッドは必ずすべて実装する**: 一つでも欠けるとコンパイルエラーになる
- **`CallByName` はリフレクション**: プロパティ名をタイプミスしても実行時までエラーにならない。定数を使う等で保護する
- **前処理・後処理の共通化**: `OnRecord` / `OnBreak` に加え `OnStart` (最初のレコード前) を加えると初期化処理もカプセル化できる

---

## まとめ: VBA における重複コードへの処方箋

VBA にはクロージャがないため、「ループの構造は同じだが中身の処理だけ違う」コードを書くとき、
開発者はしばしば似たようなループを何度もコピーして微妙に書き変えるという対応をとりがちである。
これはループに限らず、「前処理と後処理は共通で、途中の計算や出力だけが異なる」パターン全般に当てはまる。

こうした重複に対して、本稿で示した **インターフェース + クラス** による戦略パターンは有効な手法となる。
「変わらない構造」をプロシージャやモジュールに切り出し、「変わる部分」をインターフェースの実装として差し替える設計にすることで、
同じロジックを何度も書くことなく、処理の種類だけを追加できる。

ループ処理以外にも、たとえば次のような場面で同じアプローチが応用できる:

- **レポート出力**: 出力先（Debug.Print / ファイル / シート書き込み）だけ差し替えたい
- **バリデーション**: チェック項目のリストは共通で、各チェックのロジックだけ違う
- **変換処理**: データの読み取り・書き込みは共通で、変換ルールだけ差し替えたい

「構造と中身を分ける」という設計の考え方は、言語の制約に関わらず重複を減らし、変更に強いコードにするための基本原則である。
VBA の場合どうしてもコード量が多く記述箇所が分散してわかりにくくなる印象だが（実際その場面の方が多い）、
実装パターンとして知っておけば必ず役に立つ場面はあるはずである。

---

## サンプルソース

このドキュメントのコードは `sample/src/vba/closure/` に実行可能なファイルとして置いてある。

| ファイル | 役割 |
|---------|------|
| `IKeyBreakBody.cls` | インターフェース（Step 1） |
| `KeyBreakModule.bas` | 汎用ループ構造 `RunKeyBreak`（Step 2） |
| `DeptSubtotalBody.cls` | 売上小計の実装クラス（Step 3） |
| `DeptCountBody.cls` | 件数カウントの実装クラス（「別のグループ処理への使い回し」） |
| `SalesRecord.cls` | テスト用データクラス（Department / Amount フィールド） |
| `ReportModule.bas` | 呼び出し側（Step 4） |
| `ClosureTest.bas` | テストドライバ（エンジン上で実行して結果を検証） |

```bash
# 売上小計レポートを実行（出力例: 営業: 300 / 開発: 450 / 総務: 80）
./vba-run.sh sample/src/vba/closure/ClosureTest.bas RunSubtotalTest

# 件数レポートを実行（出力例: 営業: 2件 / 開発: 2件 / 総務: 1件）
./vba-run.sh sample/src/vba/closure/ClosureTest.bas RunCountTest
```

`vba-run.sh` は初回または `vba-run.ts` が更新されていれば自動でビルドしてから実行する。戻り値が stdout、`Debug.Print` は stderr に出力される。

---

## 参照

- [REFACTORING_TESTING_CATALOG.md — R-10](REFACTORING_TESTING_CATALOG.md#r-10)
