# レガシー VBA リファクタリング：どこから手を付けるか

> 対象: テストのないレガシー VBA に着手したい利用者
>
> 前提: [TESTING_STRATEGY.md](TESTING_STRATEGY.md)
>
> 次に読む: [REFACTORING_GUIDE.md](REFACTORING_GUIDE.md)、手法一覧は [REFACTORING_TESTING_CATALOG.md](REFACTORING_TESTING_CATALOG.md)

## このドキュメントの目的

テストのないレガシー VBA コードにはじめて向き合うとき、「どこから手を付けるか」「何の手法を使えばよいか」の判断を助けるガイドです。

*Working Effectively with Legacy Code*（Michael Feathers）の考え方を VBA に適用した概要版です。
各手法の詳細は [`REFACTORING_TESTING_CATALOG.md`](REFACTORING_TESTING_CATALOG.md) を参照してください。

---

## 基本原則：変更する前に守る

Feathers の核心は **「変更する前にテストで守る」** という順序です。

```
❌ リファクタリング → テストを書く   （壊れたことに気づかない）
✅ テストで守る → リファクタリング   （壊れたらすぐわかる）
```

テストがないコードを変更するのは、安全網なしで綱渡りするようなものです。
VBA Runner があることで、Excel を開かずにこの安全網を素早く敷けます。

---

## Step 1: まず分析する（触る前に全体を把握する）

コードを読む前に vba-analyzer で全体像を把握します。

```bash
# ベースラインを保存（リファクタリング効果の比較用）
npx tsx test-libs/vba-analyzer.ts <対象ディレクトリ> --json > .vba-baseline.json

# 全体の構造を把握
npx tsx test-libs/vba-analyzer.ts <対象ディレクトリ> --outline
```

**見るべき指標:**

| 指標 | 意味 | 対処の優先度 |
|---|---|---|
| `lineCount` が大きい関数 | 「神 Sub」候補。責務が混在している | 高 |
| `maxNestDepth` が深い関数 | 条件・ループが複雑に絡み合っている | 高 |
| `referenceCount = 0` の Private Sub | 呼ばれていない（Dead code 候補） | 低（削除で済む） |
| `excelAccessCount` が多い関数 | Excel 依存が強い。テストに Excel が必要 | 中（分離が必要） |
| `duplicateBlocks` | 同じロジックが複数箇所に存在 | 中（統合で簡素化） |

---

## Step 2: 手を付ける順番を決める

### 優先度の判断基準

```
触らなければならない箇所から始める
    ↓
その箇所を安全に変更できる最小の足場を作る
    ↓
足場ができたら変更する
```

**最初に手を付けるべき箇所:**

1. **バグが報告されている箇所** — 直さなければならないため、先にテストを書いて守る
2. **これから機能追加する箇所** — 変更前に特性テストを書いてベースラインを確保する
3. **行数・ネスト深さが突出している関数** — リスクが高く、改善効果も大きい

**後回しにしてよい箇所:**

- 長年変更されておらず、安定して動いているコード
- Dead code（参照数 0 の Private Sub）— 削除するだけ、テスト不要

---

## Step 3: 状況別の手法選択

### 状況 A: 「この Sub が何をしているかわからない」

**手法: 特性テスト（Characterization Test）→ [`T-14`](REFACTORING_TESTING_CATALOG.md)**

仕様を理解しようとするのではなく、現在の出力をそのまま記録します。
「今のコードが何をしているか」のスナップショットを取ることが目的です。

```
1. 代表的な入力でコードを実行する
2. 出力を記録する（正しいかどうかは問わない）
3. その出力を期待値としてテストに書く
4. これでリファクタリング後の「壊れ検知器」ができた
```

---

### 状況 B: 「テストを書きたいが Excel に依存していて書けない」

**手法: 接合部の作成（Seam）+ 純粋関数の切り出し（Extract Function）→ [`R-01`](REFACTORING_TESTING_CATALOG.md), [`R-13`](REFACTORING_TESTING_CATALOG.md)**

Excel 依存（`Range`, `Sheets`, `ActiveCell` 等）とビジネスロジックを分離します。

```vb
' Before: Excel と計算が混在（テスト不可）
Sub CalcBonus()
    Dim score As Long
    score = Sheets("Data").Range("B2").Value   ' Excel 依存
    If score >= 80 Then
        Sheets("Result").Range("C2").Value = score * 0.2  ' Excel 依存
    End If
End Sub

' After: ロジックを純粋関数に切り出す（テスト可能）
Function CalcBonusAmount(ByVal score As Long) As Double  ' ← これだけテストできる
    If score >= 80 Then
        CalcBonusAmount = score * 0.2
    End If
End Function

Sub CalcBonus()   ' Excel I/O は薄いラッパーに留める
    Dim score As Long
    score = Sheets("Data").Range("B2").Value
    Sheets("Result").Range("C2").Value = CalcBonusAmount(score)
End Sub
```

**判断基準:** Excel オブジェクトに触れない部分（計算・判定・文字列処理）はすべて切り出せます。

---

### 状況 C: 「既存コードを変えずに新機能を追加したい」

「既存コードを変えたくない」という動機はさまざまです。
*Working Effectively with Legacy Code*（Feathers）はこの状況に対していくつかの手法を示しています。
目的に合わせて選んでください。

---

**① Sprout Method（発芽メソッド）の基本パターン → [`R-17`](REFACTORING_TESTING_CATALOG.md)**

新機能のロジックを別 Sub に書き、既存コードには呼び出しを 1 行追加するだけにします。
既存の処理本体には触れないため、デグレのリスクを最小化できます。

```vb
' 既存コード（処理本体は変更しない）
Sub ProcessOrders()
    ' ... 既存処理 ...
    Call NotifySlack(orderId)   ' ← 呼び出しを 1 行追加するだけ
End Sub

' 新機能はここに書く（既存コードには触れない）
Sub NotifySlack(ByVal orderId As String)
    ' 新機能の実装
End Sub
```

---

**② 既存処理の途中に新機能を割り込ませる → Sprout Method（発芽メソッド）[`R-17`](REFACTORING_TESTING_CATALOG.md)**

新機能を追加するとき、既存ループの中にコードを混ぜ込みたくなりがちです。
よくあるのが「ループの中に処理を追加する」パターンですが、
**新しいループ処理として後から実行する** 方が既存コードへの影響を最小にできます。

```vb
' ❌ 避けたい: 新機能を既存ループの中に混ぜ込んでいる
Sub ProcessOrders()
    Dim i As Long
    For i = 2 To LastRow
        Dim orderId As String
        orderId = Cells(i, 1).Value
        SaveOrder orderId, CalcTotal(i)
        ' ↓ 新機能が既存ループの中に絡み合っている
        If NeedsNotification(i) Then
            SendMail orderId
        End If
    Next i
End Sub
```

```vb
' ✅ 推奨: 既存ループはそのまま。新機能は Sprout Method として後から実行する
Sub ProcessOrders()
    Dim i As Long
    For i = 2 To LastRow
        SaveOrder Cells(i, 1).Value, CalcTotal(i)  ' 既存処理（変更しない）
    Next i
    Call SendNotifications()  ' ← 末尾に呼び出しを 1 行追加するだけ
End Sub

' 新機能のループはここに閉じる（既存ループとは独立してテストできる）
Sub SendNotifications()
    Dim i As Long
    For i = 2 To LastRow
        If NeedsNotification(i) Then
            SendMail Cells(i, 1).Value
        End If
    Next i
End Sub
```

ループを 2 回回すことになりますが、各ループの責務が明確になり、
新機能を既存処理から切り離してテスト・修正できるメリットの方が大きいです。

---

**③ 既存処理の前後に処理を追加する → Wrap Method（ラップメソッド）[`R-18`](REFACTORING_TESTING_CATALOG.md)**

既存 Sub をリネームして内部に隠し、同名の新 Sub でラップします。
呼び出し元を一切変えずに、前処理・後処理を追加できます。

```vb
' 既存の ProcessOrders を ProcessOrders_Core にリネーム（中身は変えない）
Sub ProcessOrders_Core()
    ' ... 既存処理（変更しない）...
End Sub

' 同名の新 Sub でラップ。呼び出し元からは変化が見えない
Sub ProcessOrders()
    Call BeforeProcess()        ' ← 前処理を追加
    Call ProcessOrders_Core()   ' ← 既存処理をそのまま呼ぶ
    Call AfterProcess()         ' ← 後処理を追加
End Sub
```

---

**④ 既存処理は変えず、呼び出し元で新旧を組み合わせる**

呼び出し元を変更できる場合は、呼び出し元で新旧の処理を並べて呼ぶだけで済みます。
実装モジュールには一切手を加えません。

```vb
' 呼び出し元（ボタンクリックなど）
Sub Button_Click()
    Call ProcessOrders()    ' ← 既存処理（変更しない）
    Call NotifySlack()      ' ← 新機能を後から追加
End Sub
```

---

**選び方の目安**

| 状況 | 手法 |
|---|---|
| 新機能が既存処理と無関係 | ① 新 Sub をただ追加 |
| 既存処理の途中に割り込みたい | ② Sprout Method |
| 既存処理の前後に処理を追加したい | ③ Wrap Method |
| 呼び出し元を変えられる | ④ 呼び出し元で並べる |

---

### 状況 D: 「深いネストと複雑な条件分岐を整理したい」

**手法: ガード節の導入 + 条件の名前付け → [`R-03`](REFACTORING_TESTING_CATALOG.md)**

まず特性テストで守ってから、内側から外側に向かって整理します。

```vb
' Before: 深いネスト（読みにくい）
If isValid Then
    If Not isExpired Then
        If amount > 0 Then
            ' 本来の処理
        End If
    End If
End If

' After: ガード節で早期リターン（読みやすい）
If Not isValid Then Exit Sub
If isExpired Then Exit Sub
If amount <= 0 Then Exit Sub
' 本来の処理
```

---

### 状況 E: 「同じようなコードが複数箇所にある」

**手法: 重複ブロックの統合 → [`R-05`](REFACTORING_TESTING_CATALOG.md)**

vba-analyzer の `duplicateBlocks` で検出した重複を統合します。
ただし、**統合前に両方の呼び出しをテストで守る**のが前提です。

---

## Step 4: 変更後に効果を測る

```bash
# ベースラインとの差分を確認
npx tsx test-libs/vba-analyzer.ts <対象ディレクトリ> --diff .vba-baseline.json
```

改善の目安:

| 指標 | 良い変化 |
|---|---|
| `lineCount` | 長大な関数が分割されて減少 |
| `maxNestDepth` | 3 以下になる |
| `duplicateBlocks` | 0 に近づく |
| テスト数 | 増える（守られる範囲が広がる） |

---

## まとめ：判断フロー

```
コードを変更する必要がある
    │
    ├─ そのコードにテストがない
    │       │
    │       ├─ 動作が理解できる → 特性テストを書く（T-14）
    │       └─ 動作が理解できない → 特性テストで現状記録（T-14）
    │
    ├─ テストを書こうとしたが Excel 依存で書けない
    │       └─ ロジックを純粋関数に切り出す（R-01, R-13）
    │
    ├─ 既存コードは変えたくない（新機能追加）
    │       ├─ 新機能が独立している → 新 Sub をただ追加
    │       ├─ 既存処理の途中に割り込む → Sprout Method（R-17）
    │       ├─ 既存処理の前後に追加する → Wrap Method（R-18）
    │       └─ 呼び出し元を変えられる → 呼び出し元で新旧を並べる
    │
    └─ テストで守れた → リファクタリングを実行する
```

**「触らなければならない箇所だけ守る」** — これが Feathers の教えの核心です。
すべてのコードをきれいにしようとしないこと。変更する箇所を最小限に絞り、その箇所だけ安全にする。

---

## 参照ドキュメント

- [`REFACTORING_TESTING_CATALOG.md`](REFACTORING_TESTING_CATALOG.md) — 手法の詳細・コード例
- [`FEATHERS_TECHNIQUES_TODO.md`](FEATHERS_TECHNIQUES_TODO.md) — Feathers 手法の収録状況一覧
- [`REFACTORING_GUIDE.md`](REFACTORING_GUIDE.md) — Excel オブジェクト分離の詳細ガイド
- [`TESTING_STRATEGY.md`](TESTING_STRATEGY.md) — テスト戦略の全体像
