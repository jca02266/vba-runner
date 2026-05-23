# テーブル駆動リファクタリング Phase 1: テーブル構造提案 実装仕様

## 概要

Phase 1 では、検出されたテーブル駆動候補から **Type 定義 + 初期化関数テンプレート**を自動生成する機能を実装します。

**実装難度**: ⭐⭐（中程度）
**期待効果**: リファクタリング前工程の 30-40% 削減
**実装期間**: 1-2週間

---

## 入力・出力仕様

### 入力

検出済みの `TableDrivenCandidate` オブジェクト：

```typescript
{
  functionName: "GetApprover",
  outerBranchCount: 5,           // department
  innerBranchCount: 3,           // amount thresholds
  totalBranchCombinations: 15,
  currentLines: 56,
  linesToSave: 20,
  confidenceScore: 98,
  // ... その他メトリクス
}
```

### 出力

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

Dim g_rules() As ApprovalRule

Sub InitializeApprovalRules()
    ReDim g_rules(4)
    
    g_rules(0).department = "Sales"
    g_rules(0).threshold1 = 50000
    g_rules(0).approver1 = "Manager"
    ' ... (プレースホルダー：ユーザーが値を入力)
End Sub
```

---

## 実装の流れ

### ステップ 1: 外側分岐の情報を抽出

**目標**: 外側分岐（department == "Sales" など）の条件値と変数名を抽出

```typescript
interface OuterBranchInfo {
    variableName: string;        // "department"
    conditionValues: string[];   // ["Sales", "Marketing", "IT", "HR", "Finance"]
    dataType: string;            // "String"
}

// 実装
function extractOuterBranchInfo(ifStatement: IfStatement): OuterBranchInfo {
    // 1. if文のチェーンを走査
    const conditions = collectIfElseChain(ifStatement);
    
    // 2. 各条件から "variable = value" パターンを抽出
    const variableName = extractVariable(conditions[0].condition);  // "department"
    const values = conditions.map(cond => extractLiteralValue(cond.condition));
    
    // 3. 値の型を推定
    const dataType = inferType(values[0]);  // "String"
    
    return { variableName, conditionValues: values, dataType };
}
```

**実装の工夫**:
- VBA の Case 文にも対応（Select Case department）
- 比較演算子（=, <>, Like）に対応

---

### ステップ 2: 内側分岐の情報を抽出

**目標**: 内側分岐（amount < 50000 など）の条件値と結果を抽出

```typescript
interface InnerBranchInfo {
    conditionVariable: string;    // "amount"
    thresholds: number[];         // [50000, 500000, 2000000]
    results: string[];            // ["Manager", "Director", "VP", "CFO"]
    comparisonOperator: string;   // "<"
}

// 実装
function extractInnerBranchInfo(consequent: Statement[]): InnerBranchInfo {
    // 1. 内側の if-else-if を走査
    const innerIfs = findIfStatementsInBlock(consequent);
    
    // 2. 各 if から "variable < value" と "result = literal" を抽出
    const thresholds: number[] = [];
    const results: string[] = [];
    
    for (const ifStmt of innerIfs) {
        thresholds.push(extractThresholdValue(ifStmt.condition));
        results.push(extractAssignmentValue(ifStmt.consequent));
    }
    
    // 3. デフォルト値を抽出（最終 else）
    const defaultResult = extractDefaultValue(innerIfs[innerIfs.length - 1].alternate);
    results.push(defaultResult);
    
    return {
        conditionVariable: "amount",
        thresholds,
        results,
        comparisonOperator: "<"
    };
}
```

---

### ステップ 3: Type 定義を生成

```typescript
function generateTypeDefinition(
    typeName: string,              // "ApprovalRule"
    outer: OuterBranchInfo,
    inner: InnerBranchInfo
): string {
    let typeCode = `Type ${typeName}\n`;
    
    // 外側分岐のキー
    typeCode += `    ${outer.variableName} As ${outer.dataType}\n`;
    
    // 内側分岐ごと（閾値 + 結果）
    for (let i = 0; i < inner.thresholds.length; i++) {
        typeCode += `    threshold${i + 1} As ${inferType(inner.thresholds[i])}\n`;
        typeCode += `    ${inner.results[0].toLowerCase()}${i + 1} As ${inferType(inner.results[i])}\n`;
    }
    
    // デフォルト結果
    typeCode += `    default${inner.results[0].substring(0, 1).toUpperCase() + inner.results[0].substring(1)} As ${inferType(inner.results[0])}\n`;
    typeCode += `End Type\n`;
    
    return typeCode;
}

// 例
// Type ApprovalRule
//     department As String
//     threshold1 As Long
//     approver1 As String
//     threshold2 As Long
//     approver2 As String
//     defaultApprover As String
// End Type
```

**型推定ロジック**:
```typescript
function inferType(value: any): string {
    if (typeof value === 'string') {
        // "Manager" → String
        return 'String';
    } else if (typeof value === 'number') {
        // 50000 → Long
        // 0.15 → Double
        return Number.isInteger(value) ? 'Long' : 'Double';
    } else if (typeof value === 'boolean') {
        return 'Boolean';
    }
    // デフォルト
    return 'Variant';
}
```

---

### ステップ 4: 初期化関数テンプレートを生成

```typescript
function generateInitializationFunction(
    arrayName: string,             // "g_rules"
    typeName: string,              // "ApprovalRule"
    outer: OuterBranchInfo,
    tableSize: number              // 5 (部門数)
): string {
    let funcCode = `Sub Initialize${typeName}s()\n`;
    funcCode += `    ReDim ${arrayName}(${tableSize - 1})\n\n`;
    
    // 各行のテンプレート
    for (let i = 0; i < tableSize; i++) {
        funcCode += `    ${arrayName}(${i}).${outer.variableName} = ${outer.conditionValues[i] ? JSON.stringify(outer.conditionValues[i]) : '"TODO"'}\n`;
        
        // 閾値と結果（プレースホルダー）
        for (let j = 1; j <= outer.innerBranchCount; j++) {
            funcCode += `    ${arrayName}(${i}).threshold${j} = TODO_VALUE\n`;
            funcCode += `    ${arrayName}(${i}).approver${j} = "TODO_APPROVER"\n`;
        }
        
        funcCode += `    ${arrayName}(${i}).defaultApprover = "TODO_DEFAULT"\n`;
        funcCode += `\n`;
    }
    
    funcCode += `End Sub\n`;
    
    return funcCode;
}
```

**生成結果の例**:
```vba
Sub InitializeApprovalRules()
    ReDim g_rules(4)

    g_rules(0).department = "Sales"
    g_rules(0).threshold1 = TODO_VALUE
    g_rules(0).approver1 = "TODO_APPROVER"
    g_rules(0).threshold2 = TODO_VALUE
    g_rules(0).approver2 = "TODO_APPROVER"
    g_rules(0).defaultApprover = "TODO_DEFAULT"

    g_rules(1).department = "Marketing"
    ' ... (同じパターン)

End Sub
```

---

### ステップ 5: Lookup 関数テンプレートを生成

```typescript
function generateLookupFunction(
    functionName: string,          // "GetApprover"
    arrayName: string,             // "g_rules"
    typeName: string,              // "ApprovalRule"
    outer: OuterBranchInfo,
    inner: InnerBranchInfo,
    returnType: string            // "String"
): string {
    let funcCode = `Function ${functionName}(${inner.conditionVariable} As Long, ${outer.variableName} As String) As ${returnType}\n`;
    funcCode += `    Dim i As Integer\n`;
    funcCode += `    Dim rule As ${typeName}\n\n`;
    
    funcCode += `    For i = LBound(${arrayName}) To UBound(${arrayName})\n`;
    funcCode += `        rule = ${arrayName}(i)\n`;
    funcCode += `        If rule.${outer.variableName} = ${outer.variableName} Then\n`;
    
    // 内側分岐をジェネレート
    for (let j = 0; j < inner.thresholds.length; j++) {
        const ifKeyword = j === 0 ? 'If' : 'ElseIf';
        funcCode += `            ${ifKeyword} ${inner.conditionVariable} < rule.threshold${j + 1} Then\n`;
        funcCode += `                ${functionName} = rule.approver${j + 1}\n`;
    }
    
    funcCode += `            Else\n`;
    funcCode += `                ${functionName} = rule.defaultApprover\n`;
    funcCode += `            End If\n`;
    funcCode += `            Exit Function\n`;
    funcCode += `        End If\n`;
    funcCode += `    Next i\n\n`;
    
    funcCode += `    ${functionName} = "Unknown"\n`;
    funcCode += `End Function\n`;
    
    return funcCode;
}
```

**生成結果の例**:
```vba
Function GetApprover(amount As Long, department As String) As String
    Dim i As Integer
    Dim rule As ApprovalRule

    For i = LBound(g_rules) To UBound(g_rules)
        rule = g_rules(i)
        If rule.department = department Then
            If amount < rule.threshold1 Then
                GetApprover = rule.approver1
            ElseIf amount < rule.threshold2 Then
                GetApprover = rule.approver2
            ElseIf amount < rule.threshold3 Then
                GetApprover = rule.approver3
            Else
                GetApprover = rule.defaultApprover
            End If
            Exit Function
        End If
    Next i

    GetApprover = "Unknown"
End Function
```

---

## IDE 統合: Quick Action フロー

### ユーザー操作フロー

```
1. ユーザーが if-else-if チェーンを選択
          ↓
2. IDE が "Refactor" → "Table-Driven Pattern" を提案
          ↓
3. クリック → TableDrivenDetector が実行
          ↓
4. 検出成功 → プレビュー表示
   
   [プレビュー]
   ┌─────────────────────────────────────┐
   │ 検出結果：テーブル駆動化を推奨      │
   │ スコア: 98/100                      │
   │                                      │
   │ □ Type ApprovalRule             │
   │ □ Sub InitializeApprovalRules() │
   │ □ Function GetApprover()        │
   │                                      │
   │ [Generate] [Edit] [Cancel]      │
   └─────────────────────────────────────┘
          ↓
5. [Generate] クリック → コード生成 + エディタに挿入
          ↓
6. ユーザーが TODO_VALUE を実際の値に置き換え
          ↓
7. テストを実行 → 検証完了
```

---

## 実装チェックリスト

### コア機能

- [ ] `OuterBranchInfo` インターフェース実装
- [ ] `InnerBranchInfo` インターフェース実装
- [ ] `extractOuterBranchInfo()` 実装
- [ ] `extractInnerBranchInfo()` 実装
- [ ] `inferType()` 実装（型推定）
- [ ] `generateTypeDefinition()` 実装
- [ ] `generateInitializationFunction()` 実装
- [ ] `generateLookupFunction()` 実装

### テスト

- [ ] 単純な 5×4 パターン（ApprovalRules）
- [ ] 異なる データ型（String, Long, Double）
- [ ] エッジケース：分岐数が少ない（2×2）
- [ ] エッジケース：分岐数が多い（10×10）

### IDE 統合

- [ ] Quick Action 登録
- [ ] プレビュー表示機能
- [ ] コード挿入機能
- [ ] ユーザー確認ダイアログ

---

## 既知の制限事項

### Phase 1 でサポート対象外

1. **複雑な条件式**
   ```vba
   ' ❌ サポート対象外
   If amount < 50000 AND region = "Tokyo" Then
   ```
   
2. **条件値が固定でない場合**
   ```vba
   ' ❌ サポート対象外（変数参照）
   If amount < maxAmount Then
   ```

3. **複雑な結果値**
   ```vba
   ' ❌ サポート対象外（計算式）
   If amount > 1000000 Then
       result = amount * 0.15
   End If
   ```

### 将来の拡張で対応予定

- Phase 2+: Dictionary ベースの動的テーブル設計
- Phase 3+: テストケース自動生成による等価性検証

---

## 実装の利点

### ユーザー側

1. **時間短縮**: Type 定義 + 初期化テンプレートの自動生成（15-20分 → 2-3分）
2. **エラー低減**: テンプレート生成によるタイポ削減
3. **構造の統一**: IDE の推奨に従えば、一貫性が保証

### 開発チーム側

1. **検出から実装への自然な流れ**: Detect → Design → Test → Refactor
2. **ビジネス ロジック抽出の迅速化**: 提案されたテンプレートをベースに開発可能
3. **品質向上**: テンプレート + 自動テスト生成（Phase 3）のセット

---

## 予想される効果

| 指標 | 効果 |
|------|------|
| **開発時間** | 30-40% 削減 |
| **エラー率** | 50-60% 削減 |
| **コード品質** | テンプレート準拠で統一 |
| **リファクタリング判断** | 信頼度スコアで客観的判定 |

---

## 次のステップ

### Phase 1 完了後

- Phase 2: テーブルデータの自動抽出（CSV/JSON エクスポート）
- Phase 3: テストケース自動生成（等価性検証）
- Phase 4: 完全リファクタリング自動実装（Dictionary 対応）
