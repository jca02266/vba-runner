# テーブル駆動リファクタリング検出ガイド

## 概要

大量の条件分岐（if-else/select case）を、データテーブルと単純なルックアップロジックに置き換えるリファクタリング手法。

**効果:**
- コード行数削減（40-60%削減が一般的）
- ルール変更時に関数コード修正が不要（データ更新のみ）
- テスト容易性向上（ルールテーブル + 単純ロジック）
- 新規ルール追加時の複雑度増加がない

---

## リファクタリング前後の比較

### Before: ApprovalRules_Before.bas

```
GetApprover 関数が 71行
- 5つの部門分岐（if-else）
- 各部門内に4つの金額帯分岐（if-else）
- 合計 20個のパターン分岐
```

**コード特徴:**
- ネストの深さ: 3段階（部門 → 金額帯 → approver）
- 重複パターン: 各部門で「threshold → approver」の構造が繰り返される
- ルール変更影響: コードを編集してリコンパイル必要

### After: ApprovalRules_After.bas

```
InitializeApprovalRules で全ルール定義
GetApprover 関数が 25行（テーブル参照ロジックのみ）
- 1つのループ + 単純な if-else
- ルール定義と実行ロジックが分離
```

**コード特徴:**
- ネストの深さ: 2段階（テーブル検索 → threshold check）
- 重複がない: ルール構造を1度だけ定義
- ルール変更影響: テーブルの値を変更するだけ

### Advanced: ApprovalRules_Advanced.bas

```
ルールを Dictionary（辞書）で動的構築
- 外部データ源（スプレッドシート、JSON、DB）からの読み込み対応
- 新規部門追加時にコード変更不要
```

---

## 検出アルゴリズム

### 1. パターン認識：「繰り返される分岐構造」

**検出対象:**
```vba
If condition1 Then
    If condition2 Then
        result = value1
    ElseIf condition3 Then
        result = value2
    ElseIf condition4 Then
        result = value3
    Else
        result = value4
    End If
ElseIf condition5 Then
    If condition2 Then
        result = value5
    ElseIf condition3 Then
        result = value6
    ...
```

**特徴パターン:**
1. **外側の分岐** と **内側の分岐** が同じ構造を繰り返している
2. **割り当てパターン** が単純 OR 恣意的修正が繰り返されている
   - 許容範囲：
     - `result = valueN`（リテラル/識別子のみ）
     - `result = baseValue * rateN`（同じ変数への単一演算）
     - `result = calculatePrice(baseValue, thresholdN)`（パラメーターだけ異なる関数呼び出し）
   - 非許容範囲：
     - `result = value; specialLog(result)`（副作用がある）
     - 各パターンで異なる計算式（恣意的な修正の繰り返し）
     - 条件によってスキップされる処理
3. **分岐の数** が3個以上（テーブル化のメリットが出る）
4. **条件式** が「比較演算」のみ（複雑な論理演算ではない）

### 2. 検出ステップ

#### Step 1: ネストされた分岐を検出
```typescript
function detectNestedBranches(stmt: Statement): BranchPattern[] {
    // 外側の if と内側の if の両方を持つ関数を検出
    // ネスト深度 >= 2 である分岐構造
}
```

#### Step 2: 分岐構造の「形状」を正規化
```typescript
// 各分岐の結果（代入値）を「形状」として抽出
// 例: Sales部 → [50000→Manager, 500000→Director, 2000000→VP, ∞→CFO]
type BranchShape = Array<{condition: Expression, result: any}>
```

#### Step 3: 形状の「重複」を検出
```typescript
function findRepeatedShapes(branches: BranchPattern[]): Shape[] {
    // 複数の分岐が同じ「形状」（比較対象は異なるが構造が同じ）を持つかチェック
    // 繰り返し度合いが N >= 3 なら テーブル化のメリットあり
}
```

#### Step 4: テーブル化の有効性を評価
```typescript
type RefactoringScore = {
    repetitionCount: number,      // 繰り返し分岐の数
    nestingDepth: number,          // ネスト深度
    conditionComplexity: number,   // 条件式の複雑さ（0-10）
    codeLinesToSave: number,       // 削減される行数
    tableRowCount: number,         // テーブルの行数（< 20 が目安）
    score: number                  // 総合スコア（0-100）
}
```

---

## 実装検出の具体例

### vba-analyzer への統合

#### 新規検出ルール: `TableDrivenCandidate`

```typescript
// vba-analyzer.ts に追加
interface TableDrivenCandidate {
    functionName: string,
    fileName: string,
    line: number,
    
    // 分岐構造の分析
    nestingDepth: number,
    outerBranchCount: number,      // 外側 if-else の数（5 = Sales/Marketing/IT/HR/Finance）
    innerBranchCount: number,      // 内側 if-else の数（4 = threshold1-4）
    totalBranchCombinations: number, // 5 * 4 = 20
    
    // コード削減効果
    currentLines: number,          // 現在のコード行数（71行）
    estimatedAfterLines: number,   // テーブル化後（25行）
    linesToSave: number,           // 削減行数（46行）
    reductionPercent: number,      // 削減率（65%）
    
    // テーブル候補の分析
    tableRows: number,             // テーブルの行数候補（5 = 部門数）
    tableColumns: number,          // テーブルの列数候補（5 = thresholds + approvers + default）
    repeatingPatternCount: number, // パターンの繰り返し数（5回）
    
    // リファクタリングの実行可能性
    canTableDrive: boolean,
    riskLevel: 'low' | 'medium' | 'high',
    recommendation: string
}
```

#### 検出ロジック（TypeScript）

```typescript
function detectTableDrivenCandidates(proc: ProcedureInfo): TableDrivenCandidate[] {
    const candidates: TableDrivenCandidate[] = [];
    
    // 1. ネストされた if-else を検出
    const branches = collectBranches(proc.body, nestingDepth = 2);
    
    for (const branch of branches) {
        // 2. 外側の分岐構造を取得
        const outerConditions = branch.outer.map(b => extractCondition(b));
        const outerCount = outerConditions.length;
        
        // 3. 内側の分岐構造を取得
        const innerStructures = branch.outer.map(b => collectBranches(b.consequent));
        const innerCount = innerStructures[0]?.length || 0;
        
        // 4. 形状の重複を検出
        const shapes = innerStructures.map(s => normalizeShape(s));
        const isRepeating = isAllShapesSame(shapes);
        
        // 5. 代入パターンを分析
        const assignments = branch.collectAssignments();
        const assignmentPatterns = analyzeAssignmentPatterns(assignments);
        
        // 許容範囲：
        // - 全て同じ形（リテラル、単一変数、同じ関数の異なるパラメーター）
        // - パラメーターが分岐条件に対応している（テーブル化時に明確にマッピング可能）
        // 非許容範囲：
        // - 各パターンで異なる複雑な計算式
        // - 副作用がある（ログ出力、キャッシュ更新など）
        // - 条件によってスキップ（一部パターンで処理が実行されない）
        const isTableDriveable = canMapToTable(assignmentPatterns);
        
        // 6. 条件が比較演算のみか確認
        const conditions = branch.collectConditions();
        const isSimpleCondition = conditions.every(c =>
            c.type === 'BinaryOp' && 
            ['<', '<=', '>', '>=', '==', '!='].includes(c.operator)
        );
        
        // 7. テーブル化の有効性を計算
        if (isRepeating && isSimpleAssignment && isSimpleCondition && outerCount >= 3) {
            const candidate: TableDrivenCandidate = {
                functionName: proc.name,
                fileName: proc.file,
                line: branch.loc?.start.line,
                
                nestingDepth: 2,
                outerBranchCount: outerCount,
                innerBranchCount: innerCount,
                totalBranchCombinations: outerCount * innerCount,
                
                currentLines: branch.loc.end.line - branch.loc.start.line,
                estimatedAfterLines: calculateEstimatedLines(outerCount, innerCount),
                linesToSave: calculateSavings(...),
                reductionPercent: calculateReductionPercent(...),
                
                tableRows: outerCount,
                tableColumns: innerCount + 2, // thresholds + approvers + default
                repeatingPatternCount: outerCount,
                
                canTableDrive: true,
                riskLevel: assessRisk(...),
                recommendation: generateRecommendation(...)
            };
            
            candidates.push(candidate);
        }
    }
    
    return candidates;
}
```

---

## 現実的な検出基準：「恣意的修正」の判定

実際のビジネスロジックでは、各パターンが微妙に異なることが多い。
テーブル駆動化が有効かどうかの判断基準：

### Case 1: 許容範囲（テーブル化に適している）

**例：承認金額**
```vba
If department = "Sales" Then
    maxAmount = 500000        ' ← テーブル化可：定数値
ElseIf department = "IT" Then
    maxAmount = 800000        ' ← テーブル化可：定数値
End If
```
✅ テーブル化: 各部門の `maxAmount` を単純にテーブル行に格納

**例：割引計算（パラメーターのみ異なる）**
```vba
If tier = "Gold" Then
    discount = basePrice * 0.15   ' ← テーブル化可：異なるのは率だけ
ElseIf tier = "Silver" Then
    discount = basePrice * 0.10
End If
```
✅ テーブル化: `{ tier: "Gold", rate: 0.15 }` のテーブルで管理

### Case 2: グレーゾーン（検出すべきだが注意が必要）

**例：複合計算（パターンは同じが式が少し異なる）**
```vba
If customer = "VIP" Then
    price = basePrice * 0.85 + loyaltyBonus     ' ← 複雑：計算式が恣意的
ElseIf customer = "Regular" Then
    price = basePrice * 0.95                    ' ← 恣意的な修正？
End If
```
⚠️ テーブル化の難しさ：
- 単なる率の違いではなく、追加処理（`+ loyaltyBonus`）がある
- 各パターンで微妙に異なる式
- **判定**: ビジネスルール変更時に式全体をリファクタリングする必要がある可能性が高い
- **推奨**: テーブル化するなら、計算ロジック自体を関数に抽出して管理

### Case 3: 非許容範囲（テーブル化に不適切）

**例：パターンごとに異なる副作用**
```vba
If status = "Approved" Then
    result = "OK"
    LogApproval(customer)           ' ← 副作用：呼び出し元で期待されない
ElseIf status = "Pending" Then
    result = "Wait"
    ' ログなし
End If
```
❌ テーブル化不適切：副作用が異なる

**例：条件によってスキップ**
```vba
If amount < 10000 Then
    approver = "Manager"
ElseIf amount < 100000 Then
    approver = "Director"
    ' この範囲では特別ルール適用なし
ElseIf IsVIP(customer) Then        ' ← 追加条件でスキップ可能
    approver = "VP"
    ApplySpecialRule()
Else
    approver = "CFO"
End If
```
❌ テーブル化不適切：条件の論理が複雑（オーバーラップ、スキップ）

**例：恣意的な修正が繰り返されている**
```vba
If region = "US" Then
    tax = basePrice * 0.08
    shipping = 15
ElseIf region = "EU" Then
    tax = basePrice * (0.16 + vatRate)        ' ← EU用の追加ロジック
    shipping = 8
    AddEUCompliance()                         ' ← EU固有の処理
ElseIf region = "Asia" Then
    tax = basePrice * 0.05
    shipping = 20
    GrantImportSubsidy()                      ' ← アジア固有の処理
End If
```
❌ テーブル化が難しい理由：
- 各地域で異なるビジネスルール（税率計算式が異なる）
- 地域固有の関数呼び出し
- 毎回コードを手で修正している兆候
- **解決策**: テーブルと戦略パターンの組み合わせ（Strategy Pattern）が適切

---

## 検出のメトリクス

### スコアリング

```
スコア = (削減行数 / 現在の行数) * 40         # コード削減度
       + (繰り返し度合い / 最大値) * 30      # パターン繰り返し度
       + (恣意的修正度合い) * -20             # マイナス：恣意的修正が多いほどスコア低下
       + (副作用がない) * 15                   # プラス：純粋関数ならボーナス
       + (ビジネス変更頻度) * 15             # プラス：ルール変更が多いほど効果大

恣意的修正度合いの計算：
- 各パターンの処理が完全に同じ: 0%（マイナスなし）
- パラメーターだけ異なる: 10%（-2点）
- 計算式が微妙に異なる: 30%（-6点）
- 副作用や追加処理がある: 50%（-10点）
- 各パターンで全く異なるロジック: 100%（-20点、テーブル化不適切）

閾値:
- スコア >= 60: 強く推奨（テーブル化のメリット大）
- スコア 40-60: 検討の余地あり（改善可能性あり）
- スコア < 40: テーブル化非推奨（Strategy Pattern や他のリファクタリング検討）
```

### リスク評価

| 要素 | Low | Medium | High |
|---|---|---|---|
| 分岐の複雑度 | < 3 段階 | 3-4 段階 | > 4 段階 |
| ビジネスルール変更頻度 | 低い（年1-2回） | 中程度（月1-2回） | 高い（週1回以上） |
| テストカバレッジ | >= 90% | 50-90% | < 50% |
| 外部データ依存 | なし | 部分的 | あり |

---

## VBA Runner での自動化提案

### 提案機能: `/detect-table-driven <file-or-dir>`

```bash
# 使い方
./node_modules/.bin/esbuild test-libs/vba-analyzer.ts \
  --bundle --outfile=dist/vba-analyzer.cjs --platform=node

node dist/vba-analyzer.cjs sample/src/vba --table-driven --json
```

**出力例（JSON）:**
```json
{
  "tableDrivenCandidates": [
    {
      "functionName": "GetApprover",
      "file": "ApprovalRules_Before.bas",
      "line": 5,
      "nestingDepth": 2,
      "outerBranchCount": 5,
      "innerBranchCount": 4,
      "currentLines": 71,
      "estimatedAfterLines": 25,
      "linesToSave": 46,
      "reductionPercent": 64.8,
      "canTableDrive": true,
      "riskLevel": "low",
      "recommendation": "強く推奨。5つの部門ルールが完全に同じ構造で繰り返されている。テーブル化により 65% のコード削減が見込める。"
    }
  ]
}
```

---

## サンプルからの学習

### 定性的評価（このサンプルから得られる洞察）

1. **分岐の繰り返しパターン**
   - 5つの部門 × 4つの金額帯 = 20のパターン
   - → テーブルなら 5行 × 4列 = 20セル（データ）で表現可能

2. **ビジネスルール変更への対応性**
   - Before: 新規部門追加時に関数に新しい if-else を追加 → コンパイル必要
   - After: テーブルに行を追加 → 再コンパイル不要

3. **テストの容易性**
   - Before: 20パターン分岐 × 組み合わせテスト = 膨大
   - After: テーブル値 + ロジックの独立テスト → シンプル

---

## 検出の限界と代替手法

### テーブル駆動が不適切な場合

以下のケースではテーブル駆動化よりも他のリファクタリング手法が有効：

| パターン | 理由 | 代替手法 |
|---|---|---|
| **副作用が多い** | 各パターンで異なる処理（ログ、状態更新） | Strategy Pattern / Chain of Responsibility |
| **恣意的修正が繰り返される** | 各パターンで計算式・ロジックが異なる | ビジネスルール抽出 / DSL化 |
| **条件がオーバーラップ** | 複数条件が同時に満たされる場合がある | 優先度付きルールエンジン |
| **計算複雑度が高い** | 各パターンで異なる複雑な計算 | 計算ロジック関数化 + テーブル（ハイブリッド） |

### ハイブリッドアプローチ

**テーブル + 関数参照**（複雑な計算がある場合）:
```vba
Type PricingRule
    tier As String
    baseFactor As Double
    calculationFunc As String       ' 関数名を参照
    bonusFunc As String            ' ボーナス計算用関数
End Type

Function ApplyPricing(price, rule As PricingRule) As Double
    Dim basePrice = price * rule.baseFactor
    ' 関数名から動的に計算ロジックを呼び出し
    ApplyPricing = CallByName(Me, rule.calculationFunc, VbMethod, basePrice)
End Function
```

**テーブル + Strategy Pattern**（処理フローが異なる場合）:
```vba
Type ApprovalStrategy
    department As String
    threshold As Long
    approver As String
    strategy As IApprovalStrategy   ' インターフェース参照
End Type
```

---

## 検出の実現可能性レベル

### Level 1: 低コスト検出（実装済み）
✅ ネストされた if-else の形状検出
✅ 基本的な削減行数推定
❓ 恣意的修正度合いの自動判定

### Level 2: 中コスト検出（要実装）
⚠️ AST 解析で各パターンの計算式比較
⚠️ 副作用検出（関数呼び出し、状態変更）
⚠️ スコアリング公式の適用

### Level 3: 高コスト検出（AI連携必須）
❌ 「恣意的な修正」の意味的判定
❌ ビジネスルールの抽出・分類
❌ 最適なリファクタリング手法の選択提案

### 推奨：AI と人間の組み合わせ

1. **VBA Runner（自動）**: 候補検出 + スコアリング
2. **Claude（AI分析）**: スコア低めの候補について
   - なぜ各パターンが異なるのか（ビジネスルール分析）
   - テーブル化が本当に適切か
   - 代替手法の提案
3. **開発者（人間）**: 最終的なリファクタリング実行判断

---

## 今後の自動化

### Phase 1（現在）
- 手動でサンプル作成（Before/After）
- vba-analyzer で `--table-driven` 検出オプション追加

### Phase 2
- AST 解析で自動検出
- 削減行数・複雑度スコアの計算

### Phase 3
- リファクタリング提案自動生成
- 「Before → After」の変換コード自動生成

### Phase 4（最終）
- AI（Claude）と連携して自動リファクタリング実行
- スナップショットテストで Before/After の互換性確認
