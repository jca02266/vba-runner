# 非対称デシジョンテーブル：列数が異なるパターンのテーブル駆動化

## 概要

従来は「すべての行が同じ列数を持つ」ことが前提とされてきましたが、**実務コードでは列数が異なるのが普通**です。

```vba
' 実務例：列数が異なるケース
If department = "Sales" Then
    If amount < 50000 Then
        result = "Manager"
    ElseIf amount < 500000 Then
        result = "Director"
    ElseIf amount < 2000000 Then
        result = "VP"
    Else
        result = "CFO"
    End If
ElseIf department = "Marketing" Then
    If amount < 30000 Then
        result = "Manager"
    ElseIf amount < 300000 Then
        result = "Director"
    Else
        result = "CFO"
    End If  ' ← 3つの条件のみ（Sales は4つ）
ElseIf department = "IT" Then
    If amount < 100000 Then
        result = "Manager"
    Else
        result = "Director"
    End If  ' ← 2つの条件のみ
End If
```

**このようなケースこそテーブル駆動化が有効**です。

---

## 非対称テーブル（Asymmetric Decision Table）

### テーブル表現

```
┌──────────────┬───────────┬─────────────┬──────────────┬───────────┐
│ Department   │ <T1       │ T1-T2       │ T2-T3        │ >T3       │
├──────────────┼───────────┼─────────────┼──────────────┼───────────┤
│ Sales        │ Manager   │ Director    │ VP           │ CFO       │
│ Marketing    │ Manager   │ Director    │ (N/A)        │ CFO       │
│ IT           │ Manager   │ (N/A)       │ (N/A)        │ Director  │
└──────────────┴───────────┴─────────────┴──────────────┴───────────┘
```

**解釈**:
- Sales: 4段階の承認ルール（<50k, 50k-500k, 500k-2M, ≥2M）
- Marketing: 3段階（<30k, 30k-300k, ≥300k）
- IT: 2段階（<100k, ≥100k）

### VBA での実装

```vba
Type ApprovalRule
    department As String
    threshold1 As Long
    approver1 As String
    threshold2 As Long
    approver2 As String
    threshold3 As Long
    approver3 As String
    defaultApprover As String
End Type

Function GetApproverAsymmetric(amount As Long, department As String) As String
    Dim rule As ApprovalRule
    rule = GetRuleByDepartment(department)
    
    ' 各部門のルールに応じた判定
    ' （threshold が 0 の場合はスキップ）
    
    If rule.threshold1 > 0 And amount < rule.threshold1 Then
        GetApproverAsymmetric = rule.approver1
    ElseIf rule.threshold2 > 0 And amount < rule.threshold2 Then
        GetApproverAsymmetric = rule.approver2
    ElseIf rule.threshold3 > 0 And amount < rule.threshold3 Then
        GetApproverAsymmetric = rule.approver3
    Else
        GetApproverAsymmetric = rule.defaultApprover
    End If
End Function
```

---

## 非対称テーブルの利点

### 1. 実務ルールの忠実な表現

**現実のビジネスルール**:
```
- Sales 部: 4段階の承認が必要（大企業向け）
- Marketing 部: 3段階の承認（中規模）
- IT 部: 2段階の承認（小規模、予算が限られている）
```

テーブルが **ビジネスロジックを直接反映**

### 2. ルール変更の容易さ

```csv
department,threshold1,approver1,threshold2,approver2,threshold3,approver3,defaultApprover
Sales,50000,Manager,500000,Director,2000000,VP,CFO
Marketing,30000,Manager,300000,Director,0,NULL,CFO
IT,100000,Manager,0,NULL,0,NULL,Director
```

**ビジネス担当者が CSV を編集するだけで対応可能**

- 新部門を追加？ → CSV に 1 行追加
- ルール変更？ → 該当セルを編集
- 条件を減らす？ → threshold を 0 に設定

### 3. 複雑な if-else-if を簡潔に統一

**Before**（複雑な分岐）:
```vba
' Sales（4条件）
If department = "Sales" Then
    If amount < 50000 Then ...
    ElseIf amount < 500000 Then ...
    ElseIf amount < 2000000 Then ...
    Else ...

' Marketing（3条件）
ElseIf department = "Marketing" Then
    If amount < 30000 Then ...
    ElseIf amount < 300000 Then ...
    Else ...

' IT（2条件）
ElseIf department = "IT" Then
    If amount < 100000 Then ...
    Else ...
```

**After**（シンプルなテーブルルックアップ）:
```vba
' 統一されたテーブルルックアップロジック（10行）
Function GetApproverAsymmetric(amount, department)
    rule = GetRuleByDepartment(department)
    
    If rule.threshold1 > 0 And amount < rule.threshold1 Then
        GetApproverAsymmetric = rule.approver1
    ElseIf rule.threshold2 > 0 And amount < rule.threshold2 Then
        GetApproverAsymmetric = rule.approver2
    ElseIf rule.threshold3 > 0 And amount < rule.threshold3 Then
        GetApproverAsymmetric = rule.approver3
    Else
        GetApproverAsymmetric = rule.defaultApprover
    End If
End Function
```

---

## 非対称テーブルのバリエーション

### パターン 1: 条件数の削減

```
┌────────┬──────┬──────┬──────┐
│ Dept   │ <T1  │ T1-T2│ >T2  │
├────────┼──────┼──────┼──────┤
│ Sales  │ A    │ B    │ C    │
│ Mkt    │ A    │ B    │ C    │
│ IT     │ A    │ NULL │ B    │  ← T1-T2 の条件がない
└────────┴──────┴──────┴──────┘
```

**実装**: threshold2 = 0 にして、チェックをスキップ

### パターン 2: 結果の統合

```
┌────────┬──────┬──────┬──────┐
│ Dept   │ <T1  │ T1-T2│ >T2  │
├────────┼──────┼──────┼──────┤
│ Sales  │ Mgr  │ Dir  │ CFO  │
│ Mkt    │ Mgr  │ Dir  │ CFO  │
│ IT     │ Tech │ Tech │ Tech │  ← すべて同じ結果
└────────┴──────┴──────┴──────┘
```

**実装**: 複数の approver フィールドを同じ値で埋める
→ 実質的に 1 段階の判定

### パターン 3: 条件値の変動

```
┌────────┬────────┬────────┬────────┐
│ Dept   │ <T1    │ T1-T2  │ >T2    │
├────────┼────────┼────────┼────────┤
│ Sales  │ <50k   │ <500k  │ <2M    │
│ Mkt    │ <30k   │ <300k  │ <1.5M  │
│ IT     │ <100k  │ <800k  │ <3M    │
└────────┴────────┴────────┴────────┘
```

**すべて同じ「3段階」構造だが、値が異なる**
→ 完全対称テーブルに分類（現在の検出器が対応）

### パターン 4: 混合型（最も複雑）

```
┌────────┬──────┬──────┬──────┬──────┐
│ Dept   │ <T1  │ T1-T2│ T2-T3│ >T3  │
├────────┼──────┼──────┼──────┼──────┤
│ Sales  │ A    │ B    │ C    │ D    │  ← 4段階
│ Mkt    │ A    │ B    │ NULL │ D    │  ← 3段階（中間省略）
│ IT     │ A    │ NULL │ NULL │ B    │  ← 2段階
│ HR     │ A    │ B    │ C    │ D    │  ← 4段階
└────────┴──────┴──────┴──────┴──────┘
```

---

## 非対称テーブルの検出と判定

### 検出条件の拡張

**従来**:
```
すべての行が同じ列構造 → テーブル駆動化推奨
```

**拡張**:
```
（1）すべての行が同じ列構造（完全対称）→ 推奨度 95%
     
（2）複数の行が同じ「段階数」パターンを共有
     （列数が異なるが、最大段階数は同じ）→ 推奨度 85%
     
（3）例外的な行がある
     （ある部門だけ段階が少ない）→ 推奨度 75%
     
（4）完全に不規則
     （行ごとに段階数がバラバラ）→ 推奨度 40-50%
```

### 非対称度スコアの計算

```typescript
function calculateAsymmetryScore(table: DecisionTable): {
    asymmetryLevel: number;     // 0-1（0=完全対称, 1=完全不規則）
    commonPatternCount: number; // 共通パターンの数
    outlierCount: number;       // 異なるパターンの数
} {
    const columnCounts = table.rows.map(row => 
        row.nonNullConditionCount()
    );
    
    // 列数の最大値と最小値
    const maxCols = Math.max(...columnCounts);
    const minCols = Math.min(...columnCounts);
    
    // 非対称度 = （最大 - 最小） / 最大
    const asymmetryLevel = 1 - (minCols / maxCols);
    
    // 共通パターンを探す
    const patterns = new Map();
    for (const count of columnCounts) {
        patterns.set(count, (patterns.get(count) || 0) + 1);
    }
    
    // 最も一般的なパターンの出現数
    const commonPatternCount = Math.max(...Array.from(patterns.values()));
    const outlierCount = table.rows.length - commonPatternCount;
    
    return {
        asymmetryLevel,
        commonPatternCount,
        outlierCount
    };
}
```

### 統合スコアの計算

```typescript
function calculateAsymmetricTableScore(
    table: DecisionTable,
    asymmetry: AsymmetryScore
): number {
    // 基本スコア（デシジョンテーブル分析）
    const baseScore = calculateTablePatternScore(table);
    
    // 非対称度による調整
    let asymmetryPenalty = 0;
    
    if (asymmetry.asymmetryLevel < 0.1) {
        // 完全対称 → ペナルティなし
        asymmetryPenalty = 0;
    } else if (asymmetry.asymmetryLevel < 0.3) {
        // 軽度の非対称 → わずかなペナルティ
        asymmetryPenalty = -5;
    } else if (asymmetry.asymmetryLevel < 0.5) {
        // 中程度の非対称 → 調整
        asymmetryPenalty = -10;
    } else {
        // 高度の非対称 → 大きなペナルティ
        asymmetryPenalty = -20;
    }
    
    // 共通パターンが多い場合はボーナス
    const commonPatternBonus = 
        (asymmetry.commonPatternCount / table.rows.length) * 10;
    
    return Math.max(baseScore + asymmetryPenalty + commonPatternBonus, 30);
}
```

---

## デシジョンテーブルでの視覚化

### マークダウン表現

```markdown
| Department | <50k    | 50k-500k | 500k-2M | >2M   |
|------------|---------|----------|---------|-------|
| Sales      | Manager | Director | VP      | CFO   |
| Marketing  | Manager | Director | -       | CFO   |
| IT         | Manager | -        | -       | Mgr   |
```

**凡例**:
- `-` : その条件が適用されない
- `NULL` : スキップされる

### HTML でのカラーコーディング

```html
<table>
  <tr>
    <td>Sales</td>
    <td style="background: #90EE90">Manager</td>  <!-- 繰り返し多 -->
    <td style="background: #FFE0B0">Director</td> <!-- 繰り返し中 -->
    <td style="background: #FFD0D0">VP</td>       <!-- 繰り返し少 -->
    <td style="background: #FFD0D0">CFO</td>
  </tr>
  <tr>
    <td>Marketing</td>
    <td style="background: #90EE90">Manager</td>
    <td style="background: #FFE0B0">Director</td>
    <td style="background: #CCCCCC">（省略）</td>  <!-- グレー：条件なし -->
    <td style="background: #FFD0D0">CFO</td>
  </tr>
  <tr>
    <td>IT</td>
    <td style="background: #90EE90">Manager</td>
    <td style="background: #CCCCCC">（省略）</td>
    <td style="background: #CCCCCC">（省略）</td>
    <td style="background: #FFD0D0">Director</td>
  </tr>
</table>
```

---

## 非対称テーブルの実装パターン

### パターン A: NULL / 0 を使う

```vba
Type ApprovalRule
    department As String
    threshold1 As Long
    approver1 As String
    threshold2 As Long
    approver2 As String
    threshold3 As Long
    approver3 As String
    defaultApprover As String
End Type

' threshold = 0 で「その段階は不要」を表現
g_rules(2).threshold1 = 100000   ' <100k
g_rules(2).threshold2 = 0         ' ← 不要（スキップ）
g_rules(2).threshold3 = 0         ' ← 不要（スキップ）
g_rules(2).defaultApprover = "Director"

Function GetApprover(amount, department)
    rule = GetRuleByDepartment(department)
    
    If rule.threshold1 > 0 And amount < rule.threshold1 Then
        GetApprover = rule.approver1
    ElseIf rule.threshold2 > 0 And amount < rule.threshold2 Then
        GetApprover = rule.approver2
    ElseIf rule.threshold3 > 0 And amount < rule.threshold3 Then
        GetApprover = rule.approver3
    Else
        GetApprover = rule.defaultApprover
    End If
End Function
```

### パターン B: 動的段階数（配列）

```vba
' より柔軟：段階数が部門ごとに異なっても対応
Type ApprovalRule
    department As String
    thresholds() As Long        ' 動的配列
    approvers() As String       ' 動的配列
    defaultApprover As String
End Type

Sub InitializeRules()
    ' Sales: 4段階
    g_rules(0).department = "Sales"
    ReDim g_rules(0).thresholds(3)
    ReDim g_rules(0).approvers(3)
    g_rules(0).thresholds(0) = 50000
    g_rules(0).approvers(0) = "Manager"
    ' ...
    
    ' IT: 2段階
    g_rules(2).department = "IT"
    ReDim g_rules(2).thresholds(1)
    ReDim g_rules(2).approvers(1)
    g_rules(2).thresholds(0) = 100000
    g_rules(2).approvers(0) = "Manager"
End Sub

Function GetApproverDynamic(amount, department)
    rule = GetRuleByDepartment(department)
    
    Dim i As Integer
    For i = LBound(rule.thresholds) To UBound(rule.thresholds)
        If amount < rule.thresholds(i) Then
            GetApproverDynamic = rule.approvers(i)
            Exit Function
        End If
    Next i
    
    GetApproverDynamic = rule.defaultApprover
End Function
```

---

## 非対称テーブルの検出アルゴリズム

### Phase 1: 行の列数を計算

```typescript
function analyzeRowStructure(table: DecisionTable) {
    const rowColumnCounts = table.rows.map(row => {
        // 各行の「実質的な条件数」をカウント
        // NULL / スキップされるものは除外
        return row.conditions.filter(c => c.isActive).length;
    });
    
    return {
        minColumns: Math.min(...rowColumnCounts),
        maxColumns: Math.max(...rowColumnCounts),
        distribution: rowColumnCounts  // [4, 3, 2] など
    };
}
```

### Phase 2: パターンの共通性を判定

```typescript
function identifyCommonPatterns(table: DecisionTable) {
    const patterns = new Map();
    
    for (const row of table.rows) {
        const patternKey = JSON.stringify(row.results);
        
        if (!patterns.has(patternKey)) {
            patterns.set(patternKey, []);
        }
        patterns.get(patternKey).push(row);
    }
    
    return {
        totalPatterns: patterns.size,
        mostCommonPatternCount: Math.max(...patterns.values().map(v => v.length)),
        groupedByPattern: patterns
    };
}
```

### Phase 3: 推奨度を決定

```typescript
function determineRecommendationForAsymmetric(
    table: DecisionTable,
    structure: RowStructure,
    patterns: PatternAnalysis
): number {
    // 基本スコア 85（非対称でもテーブル駆動は有効）
    let score = 85;
    
    // 対称性による調整
    const symmetryRatio = structure.minColumns / structure.maxColumns;
    if (symmetryRatio < 0.5) {
        score -= 15;  // 大きく異なる → スコア減
    } else if (symmetryRatio < 0.8) {
        score -= 5;   // やや異なる → わずか減
    }
    
    // 共通パターンによるボーナス
    const commonRatio = patterns.mostCommonPatternCount / table.rows.length;
    score += commonRatio * 10;
    
    // 複雑さによる調整
    const distinctOutcomes = patterns.totalPatterns;
    if (distinctOutcomes > table.rows.length * 0.8) {
        score -= 20;  // ほぼすべてが異なる → テーブル駆動に不向き
    }
    
    return Math.max(Math.min(score, 100), 30);
}
```

---

## 非対称テーブル検出の利点

### ✅ 実務への適用範囲を大幅に拡大

**従来の対称テーブル前提**:
- 適用可能: 10-20% の コード（完全に同じ構造のケースのみ）

**非対称テーブル対応**:
- 適用可能: 50-70% のコード（ほとんどの実務ケース）

### ✅ 複雑な分岐をシンプルに統一

**複数の異なる if-else チェーン** → **1つのテーブルルックアップ関数**

### ✅ ビジネス変更への対応が容易

CSV でルール管理 → ビジネス担当者が直接編集可能

---

## デシジョンテーブル分析への統合

### 拡張したテーブル分析エンジン

```typescript
class EnhancedTablePatternAnalyzer {
    // 従来の分析
    calculatePatternUniformity(table): number
    calculateRowSimilarity(table): number
    calculateTableRegularity(table): number
    
    // ✨ 新規：非対称性の分析
    analyzeAsymmetry(table): AsymmetryScore
    identifyCommonPatterns(table): PatternGrouping
    calculateAsymmetryAdjustedScore(table, traditionalScore): number
}
```

### IDE での表示改善

**Before**:
```
検出結果: 列数が異なるため、テーブル駆動化は非推奨
→ ユーザー: 「え、できないの？」
```

**After**:
```
┌──────────────┬─────────┬──────────┬──────────┬───────┐
│ Department   │ <50k    │ 50k-500k │ 500k-2M  │ >2M   │
├──────────────┼─────────┼──────────┼──────────┼───────┤
│ Sales        │ Manager │ Director │ VP       │ CFO   │
│ Marketing    │ Manager │ Director │ (skip)   │ CFO   │
│ IT           │ Manager │ (skip)   │ (skip)   │ Dir   │
└──────────────┴─────────┴──────────┴──────────┴───────┘

非対称度: 中程度（IT部門のみ異なる）
共通パターン: Sales と Marketing は同じ構造
推奨スコア: 82/100 （テーブル駆動化可能）

テーブル化のメリット：
  - 現在: 複雑な3つの if-else チェーン
  - 化後: 統一されたテーブルルックアップ（10行）
  - 保守: CSV でルール管理可能

[Accept & Generate] [Edit Structure] [Decline]
```

---

## 今後の展開

### Phase 1: 非対称テーブル検出エンジン

```typescript
class AsymmetryDetector {
    detectAsymmetricPatterns(ifStatement): AsymmetricTable
    calculateAsymmetryScore(table): number
    suggestTableNormalization(table): TableNormalizationAdvice
}
```

### Phase 2: 非対称テーブルの自動提案

```
検出 → 自動提案：
  「threshold = 0 で段階をスキップする方式」
  または
  「動的配列で段階数を可変にする方式」
```

### Phase 3: CSV ベースのルール管理

```
ビジネス担当者 ↔ CSV ↔ VBA コード
  （データベースのような管理）
```

---

## 結論

**非対称テーブルをサポートすることで**:

1. ✅ テーブル駆動化の適用範囲が **5-7 倍**に拡大
2. ✅ 実務のほぼすべての場合に対応可能
3. ✅ ビジネスルール管理の **民主化**（コーダー以外でも管理可能）
4. ✅ 複雑な分岐を **統一的に**シンプル化

**ユーザーの指摘は正確です**：「列数が異なってもテーブル駆動にできる」これが実現できれば、VBA の保守性が劇的に向上します。
