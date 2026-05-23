# デシジョンテーブル分析によるテーブル駆動化検出

## 基本概念

**従来のアプローチ**: メトリクス + スコアリング  
→ 分岐数、複雑度、副作用などの定量的指標から判定

**新しいアプローチ**: デシジョンテーブル（意思決定表）  
→ コードの分岐をテーブル形式で視覚化し、テーブル構造から直接判定

---

## デシジョンテーブルとは

### ApprovalRules の例

**元のコード**:
```vba
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
        ' ... 同じパターン繰り返す
```

**デシジョンテーブル形式**:

| 行 | Department | amount < T1 | T1 ≤ amount < T2 | T2 ≤ amount < T3 | amount ≥ T3 | 結果値 |
|----|------------|------------|-----------------|-----------------|------------|--------|
| 1  | Sales | < 50k | 50k-500k | 500k-2M | ≥ 2M | Mgr / Dir / VP / CFO |
| 2  | Marketing | < 30k | 30k-300k | 300k-1.5M | ≥ 1.5M | Mgr / Dir / VP / CFO |
| 3  | IT | < 100k | 100k-800k | 800k-3M | ≥ 3M | Mgr / Dir / VP / CFO |
| 4  | HR | < 20k | 20k-200k | 200k-1M | ≥ 1M | Mgr / Dir / VP / CFO |
| 5  | Finance | < 10k | 10k-100k | 100k-500k | ≥ 500k | Mgr / Dir / VP / CFO |

**パターンの観察**:
- ✅ **同じ構造**: すべての行が「4つの閾値で4つの結果を決定」という同じパターン
- ✅ **行が冗長**: 5つの部門すべてが同じ「決定ロジック」を持つ
- ✅ **データの違い**: 閾値の値と部門名だけが異なる

→ **テーブル駆動化に適切**

---

## テーブル駆動化の適合性判定

### 条件 1: パターン一貫性（Pattern Uniformity）

```
すべての行が同じ「決定構造」を持つ
```

**ApprovalRules**:
```
行1:  Sales    [< 50k] [50k-500k] [500k-2M] [≥ 2M]  → [Mgr/Dir/VP/CFO]
行2:  Marketing[< 30k] [30k-300k] [300k-1.5M][≥1.5M] → [Mgr/Dir/VP/CFO]
行3:  IT       [< 100k][100k-800k][800k-3M] [≥ 3M]  → [Mgr/Dir/VP/CFO]
行4:  HR       [< 20k] [20k-200k] [200k-1M] [≥ 1M]  → [Mgr/Dir/VP/CFO]
行5:  Finance  [< 10k] [10k-100k] [100k-500k][≥ 500k]→ [Mgr/Dir/VP/CFO]

✅ すべての行が同じ「4つの条件 → 4つの結果」パターン
```

**非適切な例**:
```
行1: [condition1] [condition2] [condition3] [condition4] → [result1/result2/result3/result4]
行2: [condition1] [condition2] [condition3]             → [result1/result2/result3]  ❌ 列数が異なる
行3: [condition1] [condition2] [condition3] [condition4] → [result1]                 ❌ 結果の数が異なる
```

---

### 条件 2: 行の相似性（Row Similarity）

```
複数の行が同じ「決定パターン」を持つ
```

**メトリクス**:
```typescript
function calculateRowSimilarity(rows: DecisionTableRow[]): {
    similarity: number;      // 0-1（1 = 全く同じパターン）
    repeatingPatternCount: number;  // パターンが何回繰り返されているか
}

// 例
ApprovalRules の場合:
- 5 行すべてが同じ決定パターン
- Similarity = 1.0（100%）
- RepeatingPatternCount = 5
→ テーブル駆動化に非常に適切
```

**計算例**:
```
行1 pattern: "4 conditions → 4 results"
行2 pattern: "4 conditions → 4 results"
行3 pattern: "4 conditions → 4 results"
行4 pattern: "4 conditions → 4 results"
行5 pattern: "4 conditions → 4 results"

一致度 = 5/5 = 100%
```

---

### 条件 3: テーブルの規則性（Table Regularity）

```
テーブルの「形」が規則的（疎密度が高い）
```

**規則的（完全対称）**: 
```
| Dept | T1 | T2 | T3 | T4 |
|------|----|----|----|----|
| A    | X  | Y  | Z  | W  |
| B    | X  | Y  | Z  | W  |
| C    | X  | Y  | Z  | W  |

✅ 整形された矩形テーブル → テーブル駆動化に適切
```

**非対称（列数が異なるが対応可能）** ⭐ NEW:
```
| Dept | T1 | T2 | T3 | T4 |
|------|----|----|----|----|
| A    | X  | Y  | Z  | W  |
| B    | X  | Y  | -  | W  |  ← T3 がない
| C    | X  | -  | -  | W  |  ← T2, T3 がない

✅ 非対称だがテーブル駆動化に適切
  （threshold = 0 で段階をスキップ可能）
  → スコア 75-85/100
  詳細は ASYMMETRIC_DECISION_TABLES.md を参照
```

**不規則（テーブル駆動化不適切）**:
```
| Condition1 | Condition2 | Condition3 | Result |
|------------|------------|------------|--------|
| A          | [1-100]    | [a,b,c]    | X      |
| B          | [200-500]  | -          | Y      |  ← 条件ロジックが異なる
| C          | [600-∞]    | [d,e]      | Z      |

❌ 条件の意味が異なる → テーブル駆動化に不適切
```

---

## デシジョンテーブル生成アルゴリズム

### ステップ 1: 外側分岐を行に変換

```typescript
// 外側分岐チェーン
if (department = "Sales") { ... }
elseif (department = "Marketing") { ... }
elseif (department = "IT") { ... }

↓ 変換

| department |
|-----------|
| Sales     |
| Marketing |
| IT        |
```

### ステップ 2: 内側分岐を列に変換

```typescript
// 内側分岐チェーン
if (amount < 50000) { result = "Manager" }
elseif (amount < 500000) { result = "Director" }
elseif (amount < 2000000) { result = "VP" }
else { result = "CFO" }

↓ 変換

| Threshold | Result |
|-----------|--------|
| < 50000   | Mgr    |
| 50k-500k  | Dir    |
| 500k-2M   | VP     |
| ≥ 2M      | CFO    |
```

### ステップ 3: クロステーブルを作成

```
外側分岐の各値 × 内側分岐の各パターン

| Department | < 50k | 50k-500k | 500k-2M | ≥ 2M |
|------------|-------|----------|---------|------|
| Sales      | Mgr   | Dir      | VP      | CFO  |
| Marketing  | Mgr   | Dir      | VP      | CFO  |
| IT         | Mgr   | Dir      | VP      | CFO  |
| HR         | Mgr   | Dir      | VP      | CFO  |
| Finance    | Mgr   | Dir      | VP      | CFO  |
```

---

## パターン検出の実装

### データ構造

```typescript
interface DecisionTable {
    rows: DecisionTableRow[];
    columns: DecisionTableColumn[];
    cells: string[][];  // [row][col] = result
    
    // 分析結果
    analysis: {
        patternUniformity: number;      // 0-1: すべての行が同じ構造か
        rowSimilarity: number;           // 0-1: 行がどの程度似ているか
        tableRegularity: number;         // 0-1: テーブルが規則的か
        repeatingPatternCount: number;   // 何回同じパターンが繰り返されているか
        overallScore: number;            // 0-100: テーブル駆動化スコア
    };
}
```

### パターン一貫性の計算

```typescript
function calculatePatternUniformity(table: DecisionTable): number {
    // すべての行が同じ「列構造」を持つか
    const firstRowPattern = extractPattern(table.rows[0]);
    
    let uniformCount = 1;
    for (let i = 1; i < table.rows.length; i++) {
        const currentPattern = extractPattern(table.rows[i]);
        if (isSamePattern(firstRowPattern, currentPattern)) {
            uniformCount++;
        }
    }
    
    return uniformCount / table.rows.length;
    // 1.0 = すべての行が同じ
    // 0.5 = 半分が同じ
    // 0.0 = すべて異なる
}

function extractPattern(row: DecisionTableRow): PatternSignature {
    return {
        conditionCount: row.conditions.length,
        resultCount: row.results.length,
        conditionTypes: row.conditions.map(c => c.type),
        resultTypes: row.results.map(r => r.type),
    };
}
```

### 行の相似性の計算

```typescript
function calculateRowSimilarity(table: DecisionTable): number {
    if (table.rows.length <= 1) return 1.0;
    
    let totalSimilarity = 0;
    let comparisons = 0;
    
    // すべての行ペアを比較
    for (let i = 0; i < table.rows.length; i++) {
        for (let j = i + 1; j < table.rows.length; j++) {
            const similarity = compareRows(table.rows[i], table.rows[j]);
            totalSimilarity += similarity;
            comparisons++;
        }
    }
    
    return totalSimilarity / comparisons;
}

function compareRows(row1: DecisionTableRow, row2: DecisionTableRow): number {
    // 条件の構造が同じか
    const structureSimilarity = isSameStructure(row1, row2) ? 1.0 : 0.0;
    
    // 値は異なるが、パターンは同じか
    // e.g., 部門が異なるが、閾値と結果の論理は同じ
    const patternSimilarity = calculatePatternSimilarity(row1, row2);
    
    return (structureSimilarity * 0.5) + (patternSimilarity * 0.5);
}
```

### テーブル規則性の計算

```typescript
function calculateTableRegularity(table: DecisionTable): number {
    // テーブルが「詰まった矩形」か
    // スパース（空のセルが多い）なら規則性が低い
    
    const filledCells = table.cells.flat().filter(c => c !== null && c !== '').length;
    const totalCells = table.rows.length * table.columns.length;
    
    const density = filledCells / totalCells;
    
    // 密度が高いほど規則的（最小でも 70% 必要）
    return Math.min(density, 1.0);
}
```

### 総合スコアの計算

```typescript
function calculateTableDrivenScore(table: DecisionTable): number {
    const weights = {
        uniformity: 0.4,      // パターン一貫性が最重要
        similarity: 0.35,     // 行の相似性も重要
        regularity: 0.15,     // テーブル規則性は補助的
        minSize: 0.1,         // テーブルサイズも考慮
    };
    
    const uniformityScore = calculatePatternUniformity(table) * 100;
    const similarityScore = calculateRowSimilarity(table) * 100;
    const regularityScore = calculateTableRegularity(table) * 100;
    const minSizeScore = (table.rows.length >= 3 && table.columns.length >= 3) ? 100 : 0;
    
    return (
        uniformityScore * weights.uniformity +
        similarityScore * weights.similarity +
        regularityScore * weights.regularity +
        minSizeScore * weights.minSize
    );
}
```

---

## 視覚的検出の強み

### 従来のメトリクス方式の限界

```typescript
// 現在の方法
const score = calculateConfidenceScore(
    outerCount,      // 5
    innerCount,      // 3
    hasSimple        // true
);
// → スコア計算に「見えない」ことが多い
```

**問題**: 「なぜスコア 98/100 なのか」が直感的でない

### デシジョンテーブル方式の利点

```
視覚的に見える：
┌─────────────────────────────────┐
│ テーブル形式を見れば一目瞭然     │
├─────────────────────────────────┤
│ ✅ 行が繰り返されている        │
│ ✅ パターンが一貫している       │
│ ✅ テーブルが整形されている     │
└─────────────────────────────────┘

→ ユーザーも納得しやすい
```

---

## 実装例：テーブル形式での出力

### IDE への表示

```
✅ テーブル駆動化を推奨します

デシジョンテーブル分析結果:

┌──────────────┬─────────┬──────────┬──────────┬──────────┐
│ Department   │ < 50k   │ 50k-500k │ 500k-2M  │ ≥ 2M    │
├──────────────┼─────────┼──────────┼──────────┼──────────┤
│ Sales        │ Manager │ Director │ VP       │ CFO     │
│ Marketing    │ Manager │ Director │ VP       │ CFO     │
│ IT           │ Manager │ Director │ VP       │ CFO     │
│ HR           │ Manager │ Director │ VP       │ CFO     │
│ Finance      │ Manager │ Director │ VP       │ CFO     │
└──────────────┴─────────┴──────────┴──────────┴──────────┘

分析結果:
  パターン一貫性: 100% (5/5行が同じ構造)
  行の相似性: 100% (全行が同じパターン)
  テーブル規則性: 100% (完全な矩形テーブル)

推奨スコア: 99/100 (テーブル駆動化に非常に適切)

テーブル構造:
  行軸（カテゴリ）: department (5値)
  列軸（条件）: amount threshold (4つの範囲)
  セル値（結果）: approver (String型)

[Accept & Generate] [Edit Table] [Decline]
```

---

## 複雑なケースへの適用

### ケース 1: 不規則なテーブル

```
| Condition1 | Condition2 | Result |
|------------|------------|--------|
| A          | [1-100]    | X      |
| B          | [200-500]  | Y      |
| C          | [600-1000] | Z      |
| D          | Special    | W      |  ← 構造が異なる

パターン一貫性: 75% (3/4行が同じ)
行の相似性: 65%
→ スコア 72/100 (検討推奨)
```

### ケース 2: 複雑な複合条件

```
条件: (department = X) AND (amount > Y) AND (region = Z)

複数の外側分岐を同時に処理:
┌─────────────┬─────────────┬─────────┐
│ Dept/Region │ amount < 50k│ amount≥50k│
├─────────────┼─────────────┼─────────┤
│ Sales/Tokyo │ A           │ B         │
│ Sales/Osaka │ A           │ C         │  ← 同じ department でも異なる結果
│ Marketing.. │ D           │ E         │
└─────────────┴─────────────┴─────────┘

パターン一貫性: 100% (構造は同じ)
行の相似性: 60% (結果が異なる)
→ スコア 80/100 (テーブル駆動化可能だが慎重に)
```

---

## デシジョンテーブルの活用シーン

### 1. 検出フェーズ

```
AST → デシジョンテーブル生成 → パターン分析 → スコア計算
```

### 2. 設計フェーズ

```
「デシジョンテーブルの形」から
  ↓
「テーブル構造の提案」を生成

例: 5行×4列 → Type に 5 個のレコード + 4 列のフィールド
```

### 3. テスト生成フェーズ

```
テーブルの各セルが「1つのテストケース」

テストカバレッジ = 行数 × 列数
例: 5 × 4 = 20 テストケース自動生成
```

### 4. ビジネスロジック検証

```
テーブル形式をビジネス担当者に見せる

「この値は本当に正しいですか？」
→ CSVで値をアップデート
→ コードに反映
```

---

## 実装の段階

### Phase 1: デシジョンテーブル生成

```typescript
class DecisionTableGenerator {
    generateFromIfStatement(ifStmt: IfStatement): DecisionTable
    visualizeAsMarkdown(table: DecisionTable): string
    visualizeAsHTML(table: DecisionTable): string
}
```

### Phase 2: パターン分析

```typescript
class TablePatternAnalyzer {
    calculatePatternUniformity(table: DecisionTable): number
    calculateRowSimilarity(table: DecisionTable): number
    calculateTableRegularity(table: DecisionTable): number
    calculateScore(table: DecisionTable): number
}
```

### Phase 3: テーブル最適化提案

```typescript
class TableOptimizationProposer {
    // テーブルの形を変えて、より効率的な構造を提案
    proposeOptimalStructure(table: DecisionTable): DecisionTable
    suggestRowNormalization(table: DecisionTable): string[]
    suggestDataExtraction(table: DecisionTable): string  // CSV形式
}
```

---

## 期待される効果

### 検出精度の向上

| アプローチ | 正確性 | 理解しやすさ |
|-----------|--------|-----------|
| **従来：メトリクス + スコア** | 中 | 低（スコアの根拠が不透明） |
| **新規：デシジョンテーブル** | 高 | 高（テーブルを見れば一目瞭然） |

### ユーザー体験の向上

```
❌ 従来: 「スコア 98/100 です」
        「何でそんなに高いの？」

✅ 新規: テーブルを表示
        「ああ、同じパターンが5回繰り返されてるんだ」
        「なるほど、テーブル駆動化に向いてる」
```

### テスト生成への活用

```
デシジョンテーブル
  ↓
「各セルが 1 テストケース」
  ↓
20 テストケース自動生成（5行 × 4列）
  ↓
テーブル駆動化前後での等価性検証
```

---

## まとめ

**デシジョンテーブル分析**は：

1. ✅ **視覚的**: テーブルを見れば適合性が判る
2. ✅ **定量的**: パターン一貫性・行相似性・規則性を測定可能
3. ✅ **拡張可能**: テスト生成・ビジネス検証に活用可能
4. ✅ **ユーザーフレンドリー**: IDE での表示が分かりやすい

**推奨される実装順序**:
1. デシジョンテーブル生成（Phase 1-2 と並行）
2. 既存のメトリクス方式と統合
3. IDE での視覚化
4. テスト生成への活用（Phase 3 と統合）
