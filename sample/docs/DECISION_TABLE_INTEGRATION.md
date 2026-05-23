# デシジョンテーブル分析 × 検出器の統合設計

## 統合アーキテクチャ

現在の検出器を拡張し、**デシジョンテーブル分析**を組み込みます。

```
┌─────────────────────────────────────────────────┐
│ VBA コード                                       │
│ (if-else-if チェーン)                          │
└────────────┬────────────────────────────────────┘
             │
             ▼
┌─────────────────────────────────────────────────┐
│ AST パース (既存)                               │
│ → IfStatement ノードツリー                     │
└────────────┬────────────────────────────────────┘
             │
             ▼
┌─────────────────────────────────────────────────┐
│ 📊 デシジョンテーブル生成 (新規)               │
│   ├─ 外側分岐 → 行                            │
│   ├─ 内側分岐 → 列                            │
│   └─ 結果値  → セル値                         │
└────────────┬────────────────────────────────────┘
             │
             ▼
┌─────────────────────────────────────────────────┐
│ パターン分析 (新規 + 既存統合)                 │
│   ├─ 一貫性計算 (新規)                        │
│   ├─ 相似性計算 (新規)                        │
│   ├─ 規則性計算 (新規)                        │
│   └─ 従来のメトリクス (既存)                  │
└────────────┬────────────────────────────────────┘
             │
             ▼
┌─────────────────────────────────────────────────┐
│ 🎯 統合スコア計算                               │
│   テーブル分析 50% + 従来メトリクス 50%      │
└────────────┬────────────────────────────────────┘
             │
             ▼
┌─────────────────────────────────────────────────┐
│ IDE 表示 & ユーザー判定                        │
│   ├─ デシジョンテーブル視覚化                 │
│   ├─ パターン分析結果                         │
│   ├─ スコア表示（根拠付き）                   │
│   └─ [Accept / Edit / Decline]               │
└─────────────────────────────────────────────────┘
```

---

## 新規クラスの設計

### 1. `DecisionTable` インターフェース

```typescript
interface DecisionTableRow {
    keyValues: string[];           // 外側分岐のキー値
    conditions: Condition[];        // 内側分岐の条件
    results: string[];              // 各条件における結果値
}

interface DecisionTableColumn {
    conditionVariable: string;      // 条件の変数名（amount）
    thresholds: (number | string)[]; // 条件値
    operator: string;               // <, >, ==, etc.
}

interface DecisionTable {
    // テーブルの構造
    rows: DecisionTableRow[];
    columns: DecisionTableColumn[];
    cells: string[][];              // [row][col] = result
    
    // メタデータ
    metadata: {
        keyVariableName: string;    // department
        keyVariableType: string;    // String
        conditionVariableName: string;  // amount
        resultType: string;         // String
    };
}
```

### 2. `DecisionTableGenerator` クラス

```typescript
class DecisionTableGenerator {
    /**
     * IfStatement から DecisonTable を生成
     */
    generateFromIfStatement(
        ifStmt: IfStatement,
        originalCode: string
    ): DecisionTable {
        // Step 1: 外側分岐を走査 → 行を取得
        const outerChain = this.collectOuterChain(ifStmt);
        const rows = this.extractRows(outerChain);
        
        // Step 2: 内側分岐を走査 → 列を取得
        const columns = this.extractColumns(outerChain[0]);
        
        // Step 3: セル値を埋める
        const cells = this.extractCellValues(outerChain, columns);
        
        return {
            rows,
            columns,
            cells,
            metadata: this.extractMetadata(ifStmt)
        };
    }
    
    /**
     * テーブルをマークダウン形式で視覚化
     */
    visualizeAsMarkdown(table: DecisionTable): string {
        let md = `\n| ${table.metadata.keyVariableName}`;
        
        // ヘッダー行（条件）
        for (const col of table.columns) {
            md += ` | ${this.formatCondition(col)} |`;
        }
        md += `\n|${'---'.repeat(table.columns.length + 1)}|\n`;
        
        // データ行
        for (let i = 0; i < table.rows.length; i++) {
            md += `| ${table.rows[i].keyValues[0]}`;
            for (let j = 0; j < table.columns.length; j++) {
                md += ` | ${table.cells[i][j]} |`;
            }
            md += `\n`;
        }
        
        return md;
    }
    
    /**
     * テーブルを HTML 形式で視覚化
     */
    visualizeAsHTML(table: DecisionTable): string {
        let html = '<table border="1"><thead><tr>';
        html += `<th>${table.metadata.keyVariableName}</th>`;
        
        for (const col of table.columns) {
            html += `<th>${this.formatCondition(col)}</th>`;
        }
        html += '</tr></thead><tbody>';
        
        for (let i = 0; i < table.rows.length; i++) {
            html += '<tr>';
            html += `<td>${table.rows[i].keyValues[0]}</td>`;
            for (let j = 0; j < table.columns.length; j++) {
                const cellValue = table.cells[i][j];
                const bgColor = this.getCellColor(i, j, table);
                html += `<td style="background-color: ${bgColor}">${cellValue}</td>`;
            }
            html += '</tr>';
        }
        
        html += '</tbody></table>';
        return html;
    }
    
    private formatCondition(col: DecisionTableColumn): string {
        if (col.thresholds.length === 1) {
            return `${col.conditionVariable} ${col.operator} ${col.thresholds[0]}`;
        }
        // 複数の閾値の場合、範囲表記
        const t1 = col.thresholds[0];
        const t2 = col.thresholds[1];
        return `${t1} - ${t2}`;
    }
    
    private getCellColor(row: number, col: number, table: DecisionTable): string {
        // 同じ値が繰り返されている場合、色を付けて可視化
        const value = table.cells[row][col];
        const count = table.cells.flat().filter(v => v === value).length;
        
        // 繰り返し度が高いほど濃い色
        if (count >= table.rows.length * 0.8) {
            return '#90EE90';  // Light green: 高い繰り返し
        } else if (count >= table.rows.length * 0.5) {
            return '#FFFFE0';  // Light yellow: 中程度の繰り返し
        } else {
            return '#FFB6C6';  // Light pink: 繰り返しなし
        }
    }
}
```

### 3. `TablePatternAnalyzer` クラス

```typescript
class TablePatternAnalyzer {
    /**
     * パターン一貫性を計算（0-1）
     */
    calculatePatternUniformity(table: DecisionTable): number {
        if (table.rows.length <= 1) return 1.0;
        
        const firstRowStructure = this.extractStructure(table.rows[0]);
        let uniformCount = 1;
        
        for (let i = 1; i < table.rows.length; i++) {
            const currentStructure = this.extractStructure(table.rows[i]);
            if (this.isSameStructure(firstRowStructure, currentStructure)) {
                uniformCount++;
            }
        }
        
        return uniformCount / table.rows.length;
    }
    
    /**
     * 行の相似性を計算（0-1）
     * 「複数行が同じ決定パターンを持つ」度合い
     */
    calculateRowSimilarity(table: DecisionTable): number {
        if (table.rows.length <= 1) return 1.0;
        
        let totalSimilarity = 0;
        let pairCount = 0;
        
        for (let i = 0; i < table.rows.length; i++) {
            for (let j = i + 1; j < table.rows.length; j++) {
                const similarity = this.compareRowPatterns(
                    table.cells[i],
                    table.cells[j]
                );
                totalSimilarity += similarity;
                pairCount++;
            }
        }
        
        return pairCount > 0 ? totalSimilarity / pairCount : 0;
    }
    
    /**
     * テーブル規則性を計算（0-1）
     * テーブルが「詰まった矩形」か、スパースか
     */
    calculateTableRegularity(table: DecisionTable): number {
        const filledCells = table.cells
            .flat()
            .filter(c => c && c !== '' && c !== 'N/A')
            .length;
        
        const totalCells = table.rows.length * table.columns.length;
        const density = filledCells / totalCells;
        
        // 70% 以上の密度があれば規則的と判定
        return Math.min(density / 0.7, 1.0);
    }
    
    /**
     * 繰り返しパターンカウント
     */
    countRepeatingPatterns(table: DecisionTable): number {
        // 「同じパターンが何行繰り返されているか」を数える
        const patterns = new Map<string, number>();
        
        for (const row of table.rows) {
            const patternKey = row.results.join('|');  // 結果値を|で結合
            patterns.set(patternKey, (patterns.get(patternKey) || 0) + 1);
        }
        
        // 最も繰り返されているパターンの繰り返し数
        return Math.max(...Array.from(patterns.values()));
    }
    
    /**
     * 統合スコアを計算（0-100）
     * テーブル分析 50% + 従来メトリクス 50%
     */
    calculateIntegratedScore(
        table: DecisionTable,
        traditionalScore: number  // 既存の TableDrivenDetector スコア
    ): number {
        const uniformity = this.calculatePatternUniformity(table);
        const similarity = this.calculateRowSimilarity(table);
        const regularity = this.calculateTableRegularity(table);
        const repeating = this.countRepeatingPatterns(table);
        
        // テーブル分析スコア
        const tableScore = (
            uniformity * 0.4 +      // 一貫性が最重要
            similarity * 0.35 +     // 相似性も重要
            regularity * 0.15 +     // 規則性は補助
            (repeating >= 3 ? 0.1 : 0)  // 3回以上繰り返されているか
        ) * 100;
        
        // 統合スコア
        return (tableScore * 0.5) + (traditionalScore * 0.5);
    }
    
    private compareRowPatterns(row1: string[], row2: string[]): number {
        if (row1.length !== row2.length) return 0;
        
        const matches = row1.filter((val, idx) => val === row2[idx]).length;
        return matches / row1.length;
    }
}
```

---

## 既存検出器との統合

### 現在の `TableDrivenDetector` の拡張

```typescript
export class TableDrivenDetector {
    private analyzer = new BranchAnalyzer();
    private tableGenerator = new DecisionTableGenerator();
    private tableAnalyzer = new TablePatternAnalyzer();
    
    /**
     * テーブル駆動候補を検出（拡張版）
     */
    detect(funcName: string, body: Statement[], fileName: string): TableDrivenCandidate | null {
        // Step 1-3: 既存のロジック
        const nestedBranches = this.findNestedIfStatements(body);
        if (nestedBranches.length === 0) return null;
        
        const branch = nestedBranches[0] as IfStatement;
        
        // ✨ Step 4: デシジョンテーブル生成（新規）
        const decisionTable = this.tableGenerator.generateFromIfStatement(branch, '');
        
        // Step 5: パターン分析（既存メトリクス + テーブル分析）
        const traditionalScore = this.calculateConfidenceScore(...);
        
        // ✨ Step 6: テーブル分析を追加
        const tablePatternScore = this.tableAnalyzer.calculateIntegratedScore(
            decisionTable,
            traditionalScore
        );
        
        // ✨ Step 7: テーブル分析結果を追加
        const result = {
            ...existingCandidate,
            decisionTable,  // 新規フィールド
            tableAnalysis: {
                uniformity: this.tableAnalyzer.calculatePatternUniformity(decisionTable),
                similarity: this.tableAnalyzer.calculateRowSimilarity(decisionTable),
                regularity: this.tableAnalyzer.calculateTableRegularity(decisionTable),
                repeatingCount: this.tableAnalyzer.countRepeatingPatterns(decisionTable),
            },
            confidenceScore: tablePatternScore,  // 統合スコア
        };
        
        return result;
    }
}
```

---

## IDE 統合フロー

### 1. Quick Action 表示時の改善

**Before**:
```
検出結果: スコア 98/100、リスク Low
→ ユーザー: 「何でそんなに高いの？」
```

**After**:
```
| Department | <50k | 50k-500k | 500k-2M | >2M |
|------------|------|----------|---------|-----|
| Sales      | Mgr  | Dir      | VP      | CFO |
| Marketing  | Mgr  | Dir      | VP      | CFO |
| IT         | Mgr  | Dir      | VP      | CFO |

パターン一貫性: 100% (✅ すべての行が同じ構造)
行の相似性: 100% (✅ 全行が同じパターン繰り返し)
テーブル規則性: 100% (✅ 完全な矩形テーブル)

スコア: 99/100 (テーブル駆動化に非常に適切)
```

**ユーザー体験**:
「ああ、テーブルが完全に整形されてるんだ。テーブル駆動化に向いてる」

### 2. テーブルのカラーコーディング

```html
<table style="border-collapse: collapse">
  <tr>
    <th>Dept</th>
    <th><50k</th>
    <th>50k-500k</th>
    <th>500k-2M</th>
    <th>>2M</th>
  </tr>
  <tr>
    <td>Sales</td>
    <td style="background-color: #90EE90">Mgr</td>    <!-- 高繰り返し -->
    <td style="background-color: #90EE90">Dir</td>
    <td style="background-color: #90EE90">VP</td>
    <td style="background-color: #90EE90">CFO</td>
  </tr>
  <!-- ... -->
</table>
```

同じ値が繰り返されている箇所を色で強調→ パターン認識が容易

---

## 実装チェックリスト

### Phase 1: デシジョンテーブル基盤

- [ ] `DecisionTable` インターフェース定義
- [ ] `DecisionTableGenerator` クラス実装
- [ ] マークダウン / HTML 視覚化機能
- [ ] AST からのテーブル抽出ロジック

### Phase 2: パターン分析エンジン

- [ ] `TablePatternAnalyzer` クラス実装
- [ ] 一貫性・相似性・規則性計算
- [ ] 既存検出器との統合
- [ ] 統合スコア計算

### Phase 3: IDE 統合

- [ ] Quick Action 表示の改善
- [ ] テーブル形式のプレビュー
- [ ] カラーコーディング
- [ ] ユーザーフィードバック

---

## 期待される効果

### 検出精度

| 指標 | 従来 | 統合後 |
|------|------|--------|
| **正確性** | 98/100 | 99/100 (信頼度向上) |
| **説明可能性** | 低 | 高 (テーブルで明確) |
| **複雑パターン対応** | 限定的 | 柔軟 |

### ユーザー体験

| 項目 | 改善 |
|------|------|
| **理解度** | スコア → テーブル (直感的) |
| **検証可能性** | メトリクス値 → ビジュアル確認 |
| **設計への活用** | スコア → テーブル形状から設計提案 |

---

## サンプル実装

実装例は GitHub の以下で利用可能:
- `test-libs/table-driven-detector.ts` (統合前のベース)
- 新機能は別ファイル `test-libs/decision-table-generator.ts` として実装

---

## 次のステップ

1. **短期（1週間）**: DecisionTableGenerator 実装
2. **中期（2週間）**: TablePatternAnalyzer 統合
3. **長期（3週間）**: IDE フロントエンド統合
