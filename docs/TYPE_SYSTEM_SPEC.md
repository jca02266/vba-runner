# VBA型システム厳密化 仕様書（方式B: 環境メタデータ方式）

## 概要

本仕様は、VBAコンパイラの実行エンジンにおける型システムを厳密化するための設計を定義する。
JavaScript のプリミティブ `number` を維持しつつ、`Environment` クラスに変数の型メタデータを保持することで、`TypeName()` / `VarType()` の正確な挙動およびオーバーフローチェックを実現する。

## 設計方針

### 値の表現

数値は引き続き JavaScript の `number` プリミティブで保持する。ラッパークラスへの変更は行わない。
これにより `typeof val === 'number'` による判定、`===` 比較、算術演算子の直接使用がそのまま機能し、既存コードへの影響を最小化する。

### 型情報の保持

変数の宣言型を `Environment` クラス内の別マップに保持する。

```typescript
type VbaNumericType = 'Byte' | 'Integer' | 'Long' | 'Single' | 'Double' | 'Currency' | 'LongLong';
type VbaType = VbaNumericType | 'String' | 'Boolean' | 'Date' | 'Variant' | 'Object';

interface VbaTypeInfo {
    vbaType: VbaType;
}
```

```typescript
class Environment {
    private variables: Map<string, any> = new Map();
    private variableTypes: Map<string, VbaTypeInfo> = new Map();  // 追加
    // ...
}
```

## TypeName() の仕様

### 変数に対する TypeName

`Dim x As Integer` で宣言された変数に対して `TypeName(x)` は `"Integer"` を返す。

型情報の取得は環境のメタデータマップから行う。実装上、`TypeName()` 関数にはASTの引数式を渡す必要があるため、`TypeName` を通常の関数ではなく特殊構文として評価する。

| 宣言 | TypeName | VarType |
|------|----------|---------|
| `Dim x As Byte` | `"Byte"` | `17` |
| `Dim x As Integer` | `"Integer"` | `2` |
| `Dim x As Long` | `"Long"` | `3` |
| `Dim x As Single` | `"Single"` | `4` |
| `Dim x As Double` | `"Double"` | `5` |
| `Dim x As Currency` | `"Currency"` | `6` |
| `Dim x As String` | `"String"` | `8` |
| `Dim x As Boolean` | `"Boolean"` | `11` |
| `Dim x As Date` | `"Date"` | `7` |
| `Dim x` (Variant) | 格納値に依存 | 格納値に依存 |
| `Dim x As Object` | `"Object"` / クラス名 | `9` |

### リテラルに対する TypeName

リテラル値には型推論を適用する。

| リテラル | TypeName | 推論ルール |
|---------|----------|-----------|
| `10` | `"Integer"` | 整数で -32768..32767 の範囲内 |
| `40000` | `"Long"` | 整数で Integer 範囲外、Long 範囲内 |
| `10.5` | `"Double"` | 小数点を含む |
| `"abc"` | `"String"` | 文字列リテラル |
| `True` | `"Boolean"` | Boolean リテラル |
| `#2024-01-01#` | `"Date"` | 日付リテラル |

### 式の中間結果に対する TypeName

式の中間結果（`a + b` など）は Variant として扱われ、格納値の JavaScript 型から判定する。
型情報が Environment に登録されていない値に対しては、現在の挙動（`number` → `"Double"`）を維持する。

## オーバーフローチェックの仕様

型情報が登録されている変数に値を代入する際、`coerceToType()` で型制約をチェックする。

### 数値型の範囲

| 型 | 最小値 | 最大値 | 丸め |
|----|--------|--------|------|
| Byte | 0 | 255 | Banker's Rounding |
| Integer | -32,768 | 32,767 | Banker's Rounding |
| Long | -2,147,483,648 | 2,147,483,647 | Banker's Rounding |
| Single | ≈ -3.4E38 | ≈ 3.4E38 | `Math.fround()` |
| Double | (JS number の範囲) | (JS number の範囲) | なし |
| Currency | -922,337,203,685,477.5808 | 922,337,203,685,477.5807 | 小数4桁丸め |

### オーバーフロー時の挙動

範囲外の値を代入しようとした場合、VBA 実行時エラー 6 (Overflow) を発生させる。

```
Dim x As Integer
x = 40000  ' → Runtime Error 6: Overflow
```

### 型強制変換（Coercion）

型付き変数への代入時、値は対象型に強制変換される。

- **数値型への文字列代入**: `Val()` 相当の変換を試み、失敗すれば Type Mismatch（エラー 13）
- **String 型への数値代入**: `CStr()` 相当の暗黙変換
- **Boolean 型**: 0 は False(-1以外の数値は True)、数値以外は Type Mismatch

## パラメータの型情報

`Function Foo(x As Integer)` のように、パラメータに型注釈がある場合、呼び出し時のローカル環境にも型情報を登録する。
これにより、パラメータへの代入時にもオーバーフローチェックが有効になる。

### Parser の変更

`Parameter` インターフェースに `paramType?: string` フィールドを追加する。
現在 `parseParameter()` では `As` の後のトークンを `this.advance()` で読み捨てているが、これを保持するように変更する。

## 実装の影響範囲

### 変更が必要なファイル

1. **`src/compiler/evaluator.ts`**
   - `Environment` クラスに `variableTypes` マップと関連メソッドを追加
   - `evaluateVariableDeclaration()` で型情報を登録
   - `Environment.set()` で型チェック付き代入を実装
   - `TypeName()` / `VarType()` を環境メタデータ対応に改修
   - リテラルの型推論ロジックを追加

2. **`src/compiler/parser.ts`**
   - `Parameter` インターフェースに `paramType` フィールド追加
   - `parseParameter()` で型名を保持

### 変更が不要な箇所

- `evaluateBinaryExpression()`: 算術演算はそのまま（`number` のまま動作）
- `toVbaNumber()` / `toNumber()`: 変換ロジックは既存のまま
- `typeof val === 'number'` のチェック: すべてそのまま動作

## テスト計画

以下のテストケースを `tests/spec/type-system.test.ts` に実装する。

### TypeName テスト
- `Dim x As Integer: TypeName(x)` → `"Integer"`
- `Dim x As Long: TypeName(x)` → `"Long"`
- `Dim x As String: TypeName(x)` → `"String"`
- `TypeName(10)` → `"Integer"` （リテラル型推論）
- `TypeName(40000)` → `"Long"` （リテラル型推論）
- `TypeName(10.5)` → `"Double"` （リテラル型推論）

### VarType テスト
- `Dim x As Integer: VarType(x)` → `2`
- `Dim x As Long: VarType(x)` → `3`

### オーバーフローテスト
- `Dim x As Integer: x = 40000` → エラー 6
- `Dim x As Byte: x = 256` → エラー 6
- `Dim x As Byte: x = -1` → エラー 6
- `Dim x As Long: x = 3000000000` → エラー 6
- `Dim x As Integer: x = 100` → 正常（100）

### 型強制変換テスト
- `Dim x As Integer: x = 10.7` → `11`（Banker's Rounding）
- `Dim x As String: x = 123` → `"123"`

## 将来の拡張パス

方式Bから方式A（フルラップ方式）への移行が将来必要になった場合：
- `VbaTypeInfo` の定義はそのまま流用可能
- `VbaNumber` クラスに `VbaTypeInfo` を持たせる形で拡張
- `coerceToType()` のロジックも再利用可能
