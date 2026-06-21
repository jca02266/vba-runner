# テーブル駆動リファクタリング完全ガイド

## 概要

大量の**同じ構造の分岐**を、**Map（キー → 設定セット）+ 統一されたルックアップロジック**に置き換えるリファクタリング手法です。

**本質**: 「どのキーが来たか」で Map を引いて設定セットを取得し、その設定セットにしたがって結果を決める。分岐ロジックをコードではなくデータとして外出しする。

**効果**: 複雑な入れ子の `If-ElseIf` チェーンを統一された 10-20 行のコードに統合。ビジネスルール変更は「コード編集」ではなく「Map の値変更」で対応可能。

---

## 1. 実例：ApprovalRules

### Before

```vb
Function GetApprover(amount As Long, department As String) As String
    If department = "Sales" Then
        If amount < 50000 Then
            GetApprover = "Manager"
        ElseIf amount < 500000 Then
            GetApprover = "Director"
        ElseIf amount < 2000000 Then
            GetApprover = "VP"
        Else
            GetApprover = "CFO"
        End If
    ElseIf department = "Marketing" Then
        If amount < 30000 Then
            GetApprover = "Manager"
        ElseIf amount < 300000 Then
            GetApprover = "Director"
        Else
            GetApprover = "CFO"
        End If
    ElseIf department = "IT" Then
        If amount < 100000 Then
            GetApprover = "Manager"
        Else
            GetApprover = "Director"
        End If
    End If
End Function
```

**問題**: 入れ子の分岐が部門ごとに繰り返される、ルール変更=コード修正、テスト困難。

### After（Map ベース）

```vb
' Map 定義 — ビジネスルールをデータとして管理
'   キー: department（文字列）
'   値:   しきい値と承認者のペアのリスト + デフォルト
Function BuildApprovalRules() As Object
    Dim rules As Object: Set rules = CreateObject("Scripting.Dictionary")

    rules.Add "Sales",     DeptRules(Array(Level(50000, "Manager"), Level(500000, "Director"), Level(2000000, "VP")), "CFO")
    rules.Add "Marketing", DeptRules(Array(Level(30000, "Manager"), Level(300000, "Director")), "CFO")
    rules.Add "IT",        DeptRules(Array(Level(100000, "Manager")), "Director")

    Set BuildApprovalRules = rules
End Function

' ルックアップロジック — すべての部門で共通
Function GetApprover(amount As Long, department As String) As String
    Dim rules As Object: Set rules = BuildApprovalRules()
    If Not rules.Exists(department) Then GetApprover = "Unknown" : Exit Function

    Dim deptRules As Object: Set deptRules = rules(department)
    Dim levels As Variant:   levels = deptRules("levels")

    Dim i As Integer
    For i = LBound(levels) To UBound(levels)
        If amount < levels(i)("threshold") Then
            GetApprover = levels(i)("approver") : Exit Function
        End If
    Next i
    GetApprover = deptRules("default")
End Function

' ヘルパー
Function Level(threshold As Long, approver As String) As Object
    Dim t As Object: Set t = CreateObject("Scripting.Dictionary")
    t.Add "threshold", threshold: t.Add "approver", approver
    Set Level = t
End Function

Function DeptRules(levels As Variant, defaultApprover As String) As Object
    Dim r As Object: Set r = CreateObject("Scripting.Dictionary")
    r.Add "levels", levels: r.Add "default", defaultApprover
    Set DeptRules = r
End Function
```

**ポイント**:
- 各部門の段階数が異なっても（Sales=4, Marketing=3, IT=2）、Map の値のリスト長が変わるだけ。ルックアップロジックは共通。
- ルール変更 → `BuildApprovalRules()` のデータのみ修正
- 部門追加 → Map に 1 エントリ追加
- テスト容易（Map を差し替えてテストデータを注入可能）

---

## 2. デシジョンテーブル分析

コードの分岐をデシジョンテーブルで視覚化し、Map 構造への変換可能性を評価します。

### デシジョンテーブルの形式

- **行**: 個々の条件
- **セル**: Y（成立） / N（不成立） / -（前条件で確定済み、don't care）
- **列**: ルール（条件の組み合わせパターン）
- **最終行**: そのルールが成立したときの結果

### 各部門のデシジョンテーブル

**Sales（4ルール）**:
```
┌──────────────────────┬─────────┬──────────┬────┬─────┐
│ 条件                 │   R1    │    R2    │ R3 │  R4 │
├──────────────────────┼─────────┼──────────┼────┼─────┤
│ amount < 50,000      │    Y    │    N     │  N │   N │
│ amount < 500,000     │    -    │    Y     │  N │   N │
│ amount < 2,000,000   │    -    │    -     │  Y │   N │
├──────────────────────┼─────────┼──────────┼────┼─────┤
│ Approver（結果）     │ Manager │ Director │ VP │ CFO │
└──────────────────────┴─────────┴──────────┴────┴─────┘
```

**Marketing（3ルール）**:
```
┌──────────────────────┬─────────┬──────────┬─────┐
│ 条件                 │   R1    │    R2    │  R3 │
├──────────────────────┼─────────┼──────────┼─────┤
│ amount < 30,000      │    Y    │    N     │   N │
│ amount < 300,000     │    -    │    Y     │   N │
├──────────────────────┼─────────┼──────────┼─────┤
│ Approver（結果）     │ Manager │ Director │ CFO │
└──────────────────────┴─────────┴──────────┴─────┘
```

**IT（2ルール）**:
```
┌──────────────────────┬─────────┬──────────┐
│ 条件                 │   R1    │    R2    │
├──────────────────────┼─────────┼──────────┤
│ amount < 100,000     │    Y    │    N     │
├──────────────────────┼─────────┼──────────┤
│ Approver（結果）     │ Manager │ Director │
└──────────────────────┴─────────┴──────────┘
```

**観察**:
- ✅ 各部門は同じ「しきい値を順に評価する」構造
- ✅ 違いはしきい値の値と段階数だけ → Map のリスト長が変わるだけで対応可能
- ✅ `department` がキー、各部門の設定セット（しきい値リスト）が値

```
department（キー） → Map → 設定セット（しきい値リスト） → 結果
```

→ **テーブル駆動化に適切**

---

## 3. 検出条件とスコア

### 推奨される条件

| 条件 | 閾値 | 判定 |
|------|-----|------|
| **外側分岐数（Map キー数）** | ≥ 3 | ✅ 十分 |
| **内側分岐数（段階数）** | ≥ 2 | ✅ 十分 |
| **パターン一貫性** | ≥ 60% | ✅ 対応可能 |
| **行の相似性** | ≥ 50% | ✅ 有効 |
| **副作用** | なし | ✅ 安全 |
| **複雑な条件** | なし | ✅ 単純 |

### 3つのパターン分析指標

| 指標 | 説明 | 計算 |
|------|------|------|
| **パターン一貫性** | すべての部門が同じ決定構造を持つ | 0-1（1=すべて同じ） |
| **行の相似性** | 複数部門が同じパターンを繰り返す | 0-1（1=すべて同じ） |
| **テーブル規則性** | しきい値段階数がそろっているか | 最小段階数 / 最大段階数 |

### 統合スコア

```
統合スコア = （
    パターン一貫性 × 40% +
    行の相似性 × 35% +
    テーブル規則性 × 15% +
    最小サイズ × 10%
） × 100

スコア ≥ 70 → 強く推奨
スコア 50-69 → 検討推奨
スコア < 50 → 慎重に判断
```

### 非対称度

段階数が部門ごとに異なる場合の評価:

```
非対称度 = 1 - (最小段階数 / 最大段階数)

ApprovalRules の場合:
  Sales(3) + Marketing(2) + IT(1)
  = 1 - (1/3) ≈ 0.67 (中程度の非対称)
  → スコア 75-85/100 (テーブル駆動化可能)
```

---

## 4. 検出器の評価

`test-libs/table-driven-detector.ts` で AST から自動検出。

| ファイル | 状態 | 検出 | スコア |
|---------|------|------|--------|
| `ApprovalRules_Before.bas` | 分岐地獄 | ✅ | 98/100 |
| `ApprovalRules_After.bas` | テーブル駆動（Type版） | ❌ | - |
| `ApprovalRules_Advanced.bas` | テーブル駆動（Map版） | ❌ | - |

**精度**: 100%（偽陽性なし、検出漏れなし）

---

## 5. 適用範囲とトレードオフ

### 期待される効果

| 項目 | 改善 |
|------|------|
| **行数削減** | 30-70% |
| **複雑度** | ネスト深度 -1段階 |
| **保守性** | ルール変更 = Map の値変更のみ |
| **テスト化** | Map を差し替えてテストデータ注入可能 |

### 適用可否

| パターン | 適用可否 | 備考 |
|---------|---------|------|
| 完全対称（全部門で同じ段階数） | ✅ | 最も適切 |
| 軽度・中程度の非対称（段階数が異なる） | ✅ | Map のリスト長で吸収 |
| 計算式が値（`amount * 0.15` 等） | ❌ | Strategy Pattern を検討 |
| 条件が AND/OR の組み合わせ | ❌ | Strategy Pattern を検討 |
| 分岐数が少ない（≤ 2-3個） | ❌ | 効果が薄い |

### FAQ：ビジネス担当者が直接編集できるか？

CSV で管理し、Map に読み込めば可能：

```csv
department,threshold,approver
Sales,50000,Manager
Sales,500000,Director
Sales,2000000,VP
Marketing,30000,Manager
Marketing,300000,Director
```

ビジネス担当者が CSV を編集 → コード側で Map に読み込んで自動反映。

---

## 6. 実装ロードマップ

| Phase | 期間 | 内容 |
|-------|------|------|
| **Phase 1** | 1-2週間 | IDE が `if-else-if` を検出 → デシジョンテーブル視覚化 → Map 初期化コード自動生成（手作業 15-20 分 → 2-3 分） |
| **Phase 2** | 2-3週間 | 検出結果から CSV/JSON 自動抽出 → ビジネス担当者が検証・編集 → コードに自動反映 |
| **Phase 3** | 3-4週間 | Map エントリ数 × 段階数のテストケース自動生成、リファクタリング前後の等価性検証 |
| **Phase 4** | 4-6週間（研究） | 複雑なパターン・他の言語への拡張 |

---

## 7. 関連ファイル

**実装**:
- `test-libs/table-driven-detector.ts` — 検出エンジン

**サンプル**:
- `sample/src/vba/ApprovalRules_Before.bas` — リファクタリング前
- `sample/src/vba/ApprovalRules_After.bas` — リファクタリング後（Type版）
- `sample/src/vba/ApprovalRules_Advanced.bas` — リファクタリング後（Map版）

**テスト**:
- `sample/tests/ts/table-driven-detector.eval.test.ts`
- `sample/tests/ts/table-driven-detector-multi.test.ts`

**関連ドキュメント**:
- `docs/REFACTORING_GUIDE.md` パターン 5
- `docs/REFACTORING_TESTING_CATALOG.md` R-11
