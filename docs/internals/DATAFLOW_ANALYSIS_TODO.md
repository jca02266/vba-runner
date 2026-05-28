# データフロー解析 実装計画

## 1. 目的・背景

現在の静的解析（`option-explicit-checker.ts`、`symbol-table.ts`）は変数名の存在チェックに留まり、**型情報を追跡しない**。  
そのため以下が実現できていない。

| 未実現の機能 | 具体例 |
|---|---|
| Range 変数経由のセルアクセス検出 (LSP) | `Dim rng As Range` → `rng.Item(3, 5)` を Range アクセスと認識 ✅ 完了 |
| 定数識別子の値解決 (table-driven-detector) | `amount < SMALL_BUDGET` の `SMALL_BUDGET` が定数か変数か不明 ✅ 完了 |
| 静的 vs 動的名前解決の区別 | 呼び出し先が Sub/Function/変数かを静的に判定 |
| 将来的な型不一致警告 | `String` 変数に `Long` を代入した際の警告 |
| **Range メンバー補完** | `rng.` 入力時に `Value`/`Cells`/`Item`/`Offset` 等を補完（未着手） |

データフロー解析とは、「変数に対してどんな型・値が流れ込んでいるか」をコードを実行せずに追跡する静的解析技法。

---

## 2. ゴール定義

### Phase 1 ― 宣言ベース型環境の構築 ✅ 完了
**目標**: `Dim x As T` / `Const c = val` / パラメーター宣言から型マップを生成する。  
**成果物**: `TypeEnvironment` — スコープ別の `Map<varName, TypeInfo>`

```
TypeInfo = {
  kind:        'variable' | 'const' | 'parameter' | 'function'
  declaredType?: string        // 'Range', 'Long', 'String', …
  constValue?:  number|string  // Const の場合の値
  isArray:      boolean
}
```

**利用側**:
- table-driven-detector の `isNumericThresholdCondition` — `SMALL_BUDGET` が Const なら値を参照
- LSP completion — `rng.` 入力時に Range のメンバー補完

---

### Phase 2 ― 代入フロー追跡 ✅ 完了
**目標**: `Set x = expr` / `x = expr` で変数の型を絞り込む。

```
Set rng = ws.Range("A1")  →  rng の型を 'Range' と推定
x = 42                    →  x の型を 'Long/Variant(Numeric)' と推定
```

**成果物**: `FlowEnvironment` — 各代入点での型スナップショット

---

### Phase 3 ― 使用点での型照会 ✅ 完了（Range アクセス検出として実装）
**目標**: 識別子の使用箇所で `TypeEnvironment` + `FlowEnvironment` を参照し、型を返す。

```
rng.Item(3, 5)  →  rng の型 = 'Range'  →  Range のメソッド呼び出しと認識
```

**利用側**:
- LSP 診断: `rng(row, col)` が Range アクセスかどうかの判定 ✅ 実装済み（`source: 'vba-dataflow'`, severity: Hint）
- table-driven-detector: `amount < SMALL_BUDGET` の右辺が数値定数かどうか判定 ✅ 実装済み
- LSP 補完: `rng.` 入力時に Range メンバー一覧を補完 ★未着手
  - 必要な作業①: `completion-provider.ts` の `extractPrefix` を `obj.prefix` 形式に対応させる
  - 必要な作業②: 型名 → メンバー一覧の辞書を追加（`KNOWN_PROPERTY_RETURN_TYPES` の逆引き）

---

### Phase 4A ― 制御フローグラフ（CFG）構築 ★Phase 4 の前提
**目標**: プロシージャの AST を基本ブロック列 + エッジのグラフに変換する。

Phase 1–3 の `FlowEnvironment` は「線形スキャン・最後の代入を使用」という簡略モデルだった。
分岐・ループを正確に扱うには CFG が必要。

```
現在（線形）:                   CFG（Phase 4A 以降）:
  stmt1                           BB1: stmt1
  stmt2                              ↓  ↘
  stmt3                           BB2    BB3
  stmt4                              ↘  ↙
                                    BB4: stmt4
```

**成果物**: `buildCFG(proc: ProcedureDeclaration): ControlFlowGraph`

---

### Phase 4B ― 到達定義解析（Reaching Definitions）
**前提**: CFG（Phase 4A）  
**目標**: 各変数の使用箇所で「どの代入文の値が届いているか」を特定する。

```vba
x = 1           ' ← 定義 D1
If cond Then
    x = 2       ' ← 定義 D2
End If
result = x      ' ← D1 と D2 の両方が到達しうる
```

**具体的な機能**:
| 機能 | 説明 |
|---|---|
| 定義元ジャンプ | 変数の使用箇所 → 代入箇所へカーソル移動（LSP `textDocument/definition` 拡張） |
| 複数定義の検出 | 分岐によって異なる値が届く箇所を警告 |
| 変数の上書き警告 | 前の代入値が一度も読まれずに上書きされる場合（デッドストアの前段） |

**実装方式**: Gen/Kill セット per 基本ブロック + forward 不動点反復（∪）

---

### Phase 4C ― 生変数解析（Live Variable Analysis）
**前提**: CFG（Phase 4A）  
**目標**: 各時点で「この変数はこの後まだ使われるか（生きているか）」を判定する。

```vba
x = heavyCalc()   ' x がこの後 live でなければ → 無駄な計算
x = 0             ' ← x を上書き（前の値は死んでいる）
```

**具体的な機能**:
| 機能 | 説明 |
|---|---|
| デッドストア検出 | 代入後に一度も読まれない代入に警告 `x = 1` の後 `x = 2`（中間で x 未使用） |
| 未使用変数の精度向上 | 現在の「宣言されて一度も使われない」より精密な「代入値が読まれない」検出 |
| ループ内の不要代入 | ループ変数が実際に使われていない場合の検出 |

**実装方式**: Use/Def セット per 基本ブロック + backward 不動点反復（∪）

---

### Phase 4D ― 有効利用解析（Dead Store Detection）
**前提**: Phase 4C（生変数解析）  
**目標**: 代入した値が必ず後で参照されるかを判定し、参照されない代入（デッドストア）を発見する。

```vba
Sub Calc()
    Dim total As Long
    total = 0          ' ← デッドストア（次行で無条件上書き）
    total = ComputeTotal()
    ' total はここ以降使われない → total 自体もデッドストア
End Sub
```

**具体的な機能**:
| 機能 | 説明 |
|---|---|
| 無条件デッドストア | `x = a` の直後に `x = b`（中間で x を読まない） |
| 関数末尾のデッドストア | `Return` 直前の代入が戻り値以外の変数への代入 |
| コスト大な式の無駄計算 | `x = SlowFunc()` の結果 x が dead → 警告対象 |

---

### Phase 4E ― 手続き間解析（Interprocedural）
**前提**: CFG（Phase 4A）  
**目標**: 関数の戻り型を追跡し、`Set rng = GetRange()` の型を推定する。

```vba
Function GetRange() As Range
    Set GetRange = ActiveSheet.Cells
End Function

Sub Main()
    Dim r As Variant
    Set r = GetRange()   ' ← Phase 1–3 では 'Variant'、Phase 4E では 'Range'
    r(1, 1) = 99         ' ← Phase 4E で Range アクセスとして検出可能
End Sub
```

**具体的な機能**:
| 機能 | 説明 |
|---|---|
| 関数戻り型の自動推定 | 宣言型なし `Function F()` でも戻り型を推定 |
| 手続き間デッドストア | 呼び出し先が引数を使わない場合の警告 |
| ByRef 引数の追跡 | `Call Proc(x)` 後に x の型が変化しうるケースの検出 |

---

## Phase 4 全体サマリー

```
Phase 4A: CFG 構築（前提）
  ├── Phase 4B: 到達定義解析  → 定義元ジャンプ・複数定義警告
  ├── Phase 4C: 生変数解析    → デッドストア検出・未使用変数精度向上
  │     └── Phase 4D: 有効利用解析 → 無駄な代入・コスト大計算の無駄検出
  └── Phase 4E: 手続き間解析  → 関数戻り型推定・ByRef 追跡
```

複雑度が高いため Phase 1–3 の安定後に着手。

---

## 3. 現状の基盤

### 利用可能な AST ノード

| AST ノード | 型情報の在処 |
|---|---|
| `VariableDeclarator` | `.objectType?: string` — `Dim x As T` の T |
| `ProcedureDeclaration` | `.returnType?: string` / `.parameters[].type` |
| `ConstDeclaration` | `.value: Expression`（リテラルなら即値） |
| `SetStatement` | `.left`, `.right` — オブジェクト代入 |
| `AssignmentStatement` | `.left`, `.right` — 値代入 |
| `EnumDeclaration` | `.members[].name` — Enum メンバー |

### 既存の類似コード
- `option-explicit-checker.ts` — 2パス（モジュールレベル → プロシージャ）の AST 走査パターン  
- `symbol-table.ts` — スコープ別シンボルテーブルの構築パターン  
→ **Phase 1 は両者を組み合わせた拡張として実装できる**

---

## 4. 実装計画

### Step 1: `TypeEnvironment` の設計と実装
**ファイル**: `src/engine/type-environment.ts`（新規）

```typescript
export interface TypeInfo {
    kind: 'variable' | 'const' | 'parameter' | 'function' | 'enum-member';
    declaredType?: string;     // 'Range', 'Long', 'String', 'Variant', …
    constValue?: number | string | boolean;
    isArray: boolean;
}

export interface ScopeTypeMap {
    /** スコープ名（モジュールレベルは ''、プロシージャ内はプロシージャ名） */
    scopeName: string;
    vars: Map<string, TypeInfo>;  // key は小文字化した識別子名
}

export class TypeEnvironment {
    private scopes: ScopeTypeMap[] = [];

    buildFromProgram(program: Program): void { ... }

    /** 識別子の TypeInfo を返す。見つからなければ undefined */
    lookup(name: string, procName?: string): TypeInfo | undefined { ... }
}
```

---

### Step 2: `buildTypeEnvironment()` の実装
`option-explicit-checker.ts` の `collectModuleLevelDeclaredNames` に倣い、型情報も収集。

**2パス構成**:
1. **モジュールレベルパス**: `VariableDeclaration`, `ConstDeclaration`, `ProcedureDeclaration`, `EnumDeclaration` を走査
2. **プロシージャ内パス**: 各 `ProcedureDeclaration` の `parameters` + `body` を走査

**Const リテラル値の抽出**:
```typescript
// ConstDeclaration.value が NumberLiteral / StringLiteral なら即値を保存
if (decl.value.type === 'NumberLiteral') {
    typeInfo.constValue = (decl.value as NumberLiteral).value;
}
```

---

### Step 3: table-driven-detector への統合
**ファイル**: `test-libs/table-driven-detector.ts`

`isNumericThresholdCondition` での識別子解決:
```typescript
// 現在: Identifier は常に数値とみなす
const isNumericLike = (t: string) => t === 'NumberLiteral' || t === 'Identifier';

// Phase 1 以降: TypeEnvironment を参照
const isNumericLike = (node: any) => {
    if (node.type === 'NumberLiteral') return true;
    if (node.type === 'Identifier') {
        const info = this.typeEnv?.lookup(node.name);
        if (!info) return true;   // 不明な場合は数値とみなす（現行動作を維持）
        return info.kind === 'const' || isNumericType(info.declaredType);
    }
    return false;
};
```

---

### Step 4: LSP 診断への統合
**ファイル**: `src/lsp/server.ts` または新規 `src/lsp/type-checker.ts`

Range 変数経由アクセスの検出:
```typescript
// SetStatement: Set rng = ws.Range("A1") → rng の型を 'Range' として FlowEnv に記録
// MemberExpression: rng.Item(3, 5) → TypeEnv.lookup('rng') == 'Range' → 診断対象
```

---

### Step 5: Phase 2 — 代入フロー追跡
`FlowEnvironment` を実装し、`SetStatement` / `AssignmentStatement` の右辺型を推定。

**RHS 型推定の規則**:
| RHS の形 | 推定型 |
|---|---|
| `ws.Range(...)` | `'Range'` |
| `New ClassName` | `'ClassName'` |
| `CreateObject("Excel.Range")` | `'Range'`（文字列ベースの型解決） |
| 数値リテラル | `'Long'` or `'Double'` |
| 文字列リテラル | `'String'` |
| 他の変数 | その変数の FlowEnv 上の型（コピー伝播） |
| 関数呼び出し | Phase 4 まで `'Variant'` |

---

## 5. テスト計画

### Phase 1 テスト: `TypeEnvironment` 単体
**ファイル**: `tests/spec/type-environment.test.ts`（新規）

```vba
' テストコード
Const LIMIT = 100
Dim count As Long
Dim rng As Range
Function GetTotal() As Double
    Dim local As String
End Function
```

期待結果:
- `lookup('LIMIT')` → `{ kind: 'const', constValue: 100 }`
- `lookup('count')` → `{ kind: 'variable', declaredType: 'Long' }`
- `lookup('rng')` → `{ kind: 'variable', declaredType: 'Range' }`
- `lookup('GetTotal')` → `{ kind: 'function', declaredType: 'Double' }`
- `lookup('local', 'GetTotal')` → `{ kind: 'variable', declaredType: 'String' }`

### Phase 1 テスト: table-driven-detector の定数解決
`amount < LIMIT` で `LIMIT` が `Const LIMIT = 100` と同一ソースに宣言されている場合:
- `isNumericThresholdCondition` → `true`（現行と同じ動作、理由が明確になる）

### Phase 3 テスト: LSP 診断
`Dim rng As Range` + `rng.Item(3, 5)` → 診断で Range アクセスとして検出

---

## 6. スコープ外・既知の制限

| 制限 | 理由 |
|---|---|
| 別モジュールの型 | モジュール間の型伝播は Phase 4+ |
| `Object` / `Variant` 型変数の解決 | ダックタイピング → 実行時でないと確定不可 |
| 条件分岐による型ナローイング | `If TypeName(x) = "Range" Then` 等 → Phase 4+ |
| `CreateObject` の型解決 | 文字列引数に依存 → 一部対応、完全解決は困難 |
| ループ内の型変化 | `For Each x In collection` の `x` の型 |

---

## 7. 実装優先度と依存関係

```
Step 1 (TypeEnvironment 設計)
  └─ Step 2 (buildTypeEnvironment 実装)
        ├─ Step 3 (table-driven-detector 統合)  ← T-07 対応
        └─ Step 4 (LSP 診断統合)               ← Range 変数検出
              └─ Step 5 (FlowEnvironment)       ← 代入フロー
                    └─ Phase 4 (手続き間解析)
```

Step 1–3 が最小実用セット（MVP）。  
Step 4 は LSP 診断改善、Step 5 以降は精度向上。

---

## 8. 関連ファイル

| ファイル | 役割 |
|---|---|
| `src/engine/parser.ts` | AST 定義（`VariableDeclarator`, `ProcedureDeclaration` 等） |
| `src/engine/option-explicit-checker.ts` | 2パス AST 走査の先例 |
| `src/lsp/symbol-table.ts` | スコープ構造構築の先例 |
| `test-libs/table-driven-detector.ts` | Phase 1 の最初の統合対象 |
| `src/lsp/server.ts` | Phase 3 の LSP 診断統合先 |
| `src/engine/type-environment.ts` | **新規作成**（Step 1） |
