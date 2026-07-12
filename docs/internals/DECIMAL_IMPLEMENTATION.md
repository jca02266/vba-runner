# Currency / Decimal の BigInt 固定小数点実装

`Currency` 型と `Decimal` 型の内部実装解説。両型とも JS `number`（IEEE 754 double）ではなく
**BigInt 整数を内部表現とする固定小数点**で実装されている。

関連ドキュメント: [`LITERALS_AND_OPERATORS.md`](LITERALS_AND_OPERATORS.md)（リテラル型サフィックスと演算子の型規則）、
[`TYPE_SYSTEM_SPEC.md`](TYPE_SYSTEM_SPEC.md)（型変換とオーバーフローの意味論）

---

## 1. 背景

### 1.1 なぜ BigInt が必要か

Currency / Decimal は float 演算では精度が出ない:

```vba
' 実 VBA:     0.3
' float 演算: 0.30000000000000004
CCur(0.1) + CCur(0.2)
```

二つの根本的な問題がある:

1. **入力時点で誤差が混入する**: `0.1` は float で正確に表現できず、
   `0.1 * 10000 = 1000.0000000000001` のようなずれが生じる
2. **Currency の全範囲が `MAX_SAFE_INTEGER` を超える**: `±922,337,203,685,477.5807`
   を内部整数 ×10⁴ で表すと約 9.2×10¹⁸ になり、JS 安全整数上限（約 9.0×10¹⁵）を超える。
   乗算の中間値はさらに大きくなる

したがって **正確な Currency/Decimal は BigInt でしか実装できない**。

### 1.2 先行事例: LongLong

LongLong はすでに native `bigint` で実装されており、BigInt 値がエンジンを流れる前例がある。
Currency / Decimal が LongLong と異なるのは「精度低下を許容しない」点であるため、
**ラッパークラス + 明示的な演算分岐**を採用している（§4）。

---

## 2. データ表現

### 2.1 VbaCurrency（`vba-types.ts`）

```ts
export class VbaCurrency {
    readonly internal: bigint;  // 実値 × 10^4。例: 0.1 → 1000n、1.5 → 15000n

    static readonly MIN = -9223372036854775808n;  // -922,337,203,685,477.5808
    static readonly MAX =  9223372036854775807n;  //  922,337,203,685,477.5807

    constructor(internal: bigint) { /* 範囲チェック */ }
    static fromNumber(val: number): VbaCurrency { /* toFixed(4) 経由 */ }
    toString(): string { /* BigInt 除算で文字列化、末尾ゼロ除去 */ }
}
```

- スケールは **10⁻⁴ 固定**。スケールフィールドは持たない
- 範囲チェックはコンストラクターで一元化（int64 範囲 = VBA Currency の仕様範囲 ×10⁴）
- `toString()` は末尾ゼロを除去する（`CStr(CCur(1.5))` → `"1.5"` であって `"1.5000"` ではない）

### 2.2 VbaDecimal（`vba-types.ts`）

```ts
export class VbaDecimal {
    readonly mantissa: bigint;  // 符号付き。|mantissa| ≤ 2^96-1
    readonly scale: number;     // 0-28。実値 = mantissa × 10^(-scale)

    static readonly MAX_MANTISSA = 79228162514264337593543950335n; // 2^96 - 1
    static readonly MAX_SCALE = 28;

    constructor(mantissa: bigint, scale: number) { /* 範囲チェック */ }
    get value(): number { /* 後方互換ゲッター。精度は約 15-17 桁 */ }
    toString(): string { /* BigInt 演算で文字列化、末尾ゼロ除去 */ }
    static fromNumber(val: number): VbaDecimal { /* String(val) 経由 */ }
    static fromString(s: string): VbaDecimal    { /* float 非経由で直接パース */ }
}
```

- VBA 実仕様（DECIMAL 構造体: 96bit 符号なし整数 + 符号ビット + スケール 0〜28）に準拠
- **内部スケールは正規化しない**（演算後も入力時のスケールを維持）
- 表示時のみ末尾ゼロを除去（`toString()` の実装）
- 仮数が 2^96-1 を超えた場合はスケールを 1 桁ずつ下げて収める（`normalizeDecimal`）。
  スケール 0 でも収まらなければ Error 6 (Overflow)

### 2.3 `valueOf()` を定義しない理由（両クラス共通）

`valueOf()` を定義すると、JS の算術文脈に混入したとき**暗黙に float 化して静かに精度を失う**。
`valueOf()` を持たないオブジェクトが算術文脈に紛れ込むと `NaN` として顕在化するため、
**演算分岐の実装漏れをテストで検出できる**。これは意図的な設計であり、
`VbaDate` / `VbaBoolean`（`valueOf()` あり）とは方針が異なる。

数値が必要な場面（`CDbl` 等の明示的変換）はすべて専用の変換関数を通す（§5.3）。

---

## 3. 入力変換（float 誤差の回避）

「どの入力経路でも float の丸め誤差を内部表現に持ち込まない」ことが原則。

### 3.1 `CCur(number)` — float からの変換

```ts
static fromNumber(val: number): VbaCurrency {
    // 1. バンカーズ丸めを 4 桁で適用
    const rounded = vbaRound(val, 4);
    // 2. toFixed(4) で 10 進文字列化（2 進誤差を切り落とす）
    const s = rounded.toFixed(4);   // "0.1000"
    // 3. 文字列から BigInt を構築（float を経由しない）
    return new VbaCurrency(parseFixedPointString(s, 4));
}
```

`0.1 * 10000 = 1000.0000000000001` のような誤差はステップ 2〜3 で排除される。

### 3.2 `CCur(string)` / `CDec(string)` — 文字列からの直接パース

```ts
// 実装: vba-types.ts の parseFixedPointString / VbaDecimal._parse
CDec("0.1")  → mantissa=1n,   scale=1（厳密）
CDec("1.10") → mantissa=110n, scale=2（スケール保持）
CCur("0.1")  → internal=1000n
```

- 小数部が Currency の 4 桁 / Decimal の 28 桁を超える場合はバンカーズ丸め
- 指数表記（`1e-7` 等）は `expandExponential()` で平文に展開してからパース
- 数値として解釈できない文字列は Error 13、範囲超過は Error 6

### 3.3 リテラル `1.5@` — Lexer の生テキストから構築

`@` サフィックス付きリテラルは `applyLiteralTypeSuffix` で `VbaCurrency.fromNumber()` に渡す。
`parseFloat` した number を経由する実装になっているため、
15 桁超のリテラルは精度が落ちる（実用上は Currency の有効桁内なので問題なし）。

### 3.4 `CDec(number)`

number の JS 最短文字列表現 `String(val)` を §3.2 のパースに通す。
`String(0.1)` → `"0.1"` なので、ユーザーが書いたリテラルの見た目どおりの Decimal が得られる。

### 3.5 型間変換

| 変換 | 実装 |
|---|---|
| `CDec(VbaCurrency)` | `new VbaDecimal(val.internal, 4)` — scale=4 で精度保持 |
| `CCur(VbaBoolean)` | `True → -10000n`、`False → 0n` |
| `CCur(bigint)` | `new VbaCurrency(val * 10000n)` + 範囲チェック |

---

## 4. 算術演算（`evaluator.ts`）

### 4.1 分岐位置

`evaluateBinaryExpression` 内で `toVbaNumber` による float 変換の**前**に分岐する:

```ts
if (leftVal instanceof VbaDecimal || rightVal instanceof VbaDecimal) {
    return this.evaluateDecimalOp(op, leftVal, rightVal, toVbaNumber);
}
if (leftVal instanceof VbaCurrency || rightVal instanceof VbaCurrency) {
    return this.evaluateCurrencyOp(op, leftVal, rightVal, toVbaNumber);
}
```

Decimal 判定を先に置く（Decimal と他の型が混在した場合は結果を Decimal にするため）。

### 4.2 型昇格規則（MS-VBAL 準拠）

| 演算 | オペランド | 結果型 |
|---|---|---|
| `+` `-` `*` | Currency ⊕ {Byte, Integer, Long, Single, Boolean, Currency} | **Currency** |
| `+` `-` `*` | Currency ⊕ Double（小数を含む float） | **Double** |
| `+` `-` `*` `/` | Decimal ⊕ 任意の数値型 | **Decimal** |
| `/` | Currency ⊕ Currency | **Double**（`/` は常に浮動小数点除算） |
| `\` `Mod` | Currency / Decimal | 整数化してから演算、結果は Long |

### 4.3 Currency 演算（`evaluateCurrencyOp`）

```ts
case '+': return new VbaCurrency(ai + bi);
case '-': return new VbaCurrency(ai - bi);
case '*': return new VbaCurrency(bankersDivide(ai * bi, 10000n));
//  (a×10⁴)(b×10⁴) = 実値×10⁸ → 10⁴ で割ってスケールを戻す（バンカーズ丸め）
case '/': return Number(ai) / Number(bi);  // Double を返す（仕様どおり）
```

非 Currency 側のオペランドは整数なら `BigInt(val) * 10000n`、
小数を含む float なら `VbaCurrency.fromNumber(val)` を経由して `internal` を取り出す。
Double との演算は Currency 側を `Number(internal) / 10000` で float 化して Double 演算へ。

### 4.4 Decimal 演算（`evaluateDecimalOp`）

スケールが可変なため演算ごとにスケール調整が必要:

| 演算 | 実装 |
|---|---|
| `+` `-` | スケールを `max(s₁, s₂)` に整列（小さい方の mantissa を 10ⁿ 倍）してから加減算 |
| `*` | mantissa 同士を乗算、スケール = s₁ + s₂（28 超過分は `normalizeDecimal` でバンカーズ丸め縮退） |
| `/` | `bankersDivide(m₁ × 10^(28 - s₁ + s₂), m₂)`、結果スケール = 28 |
| 比較 | スケール整列後に mantissa 比較 |

**`normalizeDecimal(m, scale)`**: オーバーフロー時にスケールを 1 桁ずつ下げながら
バンカーズ丸めで仮数を縮退させる。スケール 0 でも 2^96-1 を超える場合のみ Overflow エラー。

```ts
// 例: CDec(1) / CDec(3)
//   extNum = 1n × 10^28, quotient = bankersDivide(10^28, 3n)
//   → 3333333333333333333333333333n, scale=28
//   → "0.3333333333333333333333333333"（28桁）
```

**共通ヘルパー（`vba-types.ts`）**:
- `parseFixedPointString(s, scale)`: 文字列 → BigInt（バンカーズ丸めあり）
- `bankersDivide(n, d)`: BigInt 整数除算（余りが d/2 のとき偶数丸め）
- `expandExponential(s)`: `"1e-7"` → `"0.0000001"`

---

## 5. 型変換・表示

### 5.1 `CStr` / `Debug.Print`（`coerce.ts`）

```ts
// vbaToDisplayString
if (val instanceof VbaCurrency) return val.toString();
if (val instanceof VbaDecimal)  return val.toString();
```

`toString()` は BigInt 演算で文字列化し、float を一切経由しない。末尾ゼロは除去する。

### 5.2 `Format()`

現状は **Double 経由**（`Number(internal)/10000` してから既存の `formatNumber` へ）。
15 桁を超える金額を `"#,##0.00"` 等で表示するユースケースは稀なため、当面はこの実装を維持する。
BigInt 直接対応は将来課題（§7）。

### 5.3 下位型への変換

| 変換 | 実装 |
|---|---|
| `CDbl(c)` / `CSng(c)` | `Number(c.internal) / 10000`（精度低下は仕様どおり） |
| `CLng(c)` / `CInt(c)` | `bankersDivide(c.internal, 10000n)` で BigInt のまま整数化してから `Number()` |
| `CLngLng(c)` | `bankersDivide(c.internal, 10000n)` をそのまま bigint で返す |

`vbaToNumber()`（`coerce.ts`）は `Number(val.internal) / 10000`（Currency）または
`val.value`（Decimal の getter）で float 近似値を返す。明示変換が不要な場面での後方互換用。

### 5.4 `TypeName` / `VarType`（`builtins.ts`）

```ts
// vartype
if (val instanceof VbaCurrency) return 6;
if (val instanceof VbaDecimal)  return 14;

// typename
if (val instanceof VbaCurrency) return 'Currency';
if (val instanceof VbaDecimal)  return 'Decimal';
```

AST レベルの型推論（`resolveBinaryExprNumericSubtype`）とは独立して、
実値の `instanceof` で判定するため確実に正しい型名が返る。

### 5.5 型強制（coercion）

| 場面 | 実装 |
|---|---|
| `Dim c As Currency` への代入 | `coerceToType('Currency')` → `ccur(val)` → `VbaCurrency` |
| `Dim d As Decimal` 配列要素への代入 | `coerceToDeclaredType('Decimal')` → `cdec(val)` → `VbaDecimal` |
| `Dim c As Currency` 初期値 | `new VbaCurrency(0n)` |
| `Dim d As Decimal` スカラー変数 | 初期値なし（`null`）。代入時の強制変換なし（Decimal は型マップ未登録） |

> `Dim d As Decimal` のスカラー変数は現状 coercion されない（`typeMap` に 'Decimal' が未登録）。
> `CDec()` を明示的に呼べば `VbaDecimal` が格納される。配列要素は `__vbaElementType__` 経由で強制変換される。

---

## 6. 実装箇所一覧

| ファイル | 変更箇所 |
|---|---|
| `src/engine/vba-types.ts` | `VbaCurrency` 新設、`VbaDecimal` を `mantissa`/`scale` に刷新、`parseFixedPointString` / `bankersDivide` / `expandExponential` 追加 |
| `src/engine/builtins.ts` | `ccur` BigInt 化、`cdec` を文字列直接パースに変更、`typename` / `vartype` に `VbaCurrency` 分岐追加 |
| `src/engine/evaluator.ts` | `evaluateCurrencyOp` / `evaluateDecimalOp` 新設、`evaluateBinaryExpression` に分岐追加、`evaluateUnaryExpression` に `-VbaCurrency` / `-VbaDecimal` 追加、`coerceToType('Currency')` で `VbaCurrency` 構築 |
| `src/engine/coerce.ts` | `vbaToNumber` / `vbaToDisplayString` に `VbaCurrency` / `VbaDecimal` 分岐追加 |

---

## 7. テスト

| テストファイル | 内容 |
|---|---|
| `tests/spec/currency-bigint.test.ts` | Currency BigInt 精度の主要 20 テスト（`CCur(0.1)+CCur(0.2)="0.3"` 等） |
| `tests/spec/currency-precision.test.ts` | Currency の 15 シナリオ（加減乗除・配列・関数引数・invoice 計算） |
| `tests/spec/decimal-bigint.test.ts` | Decimal 28 桁精度の 18 テスト（文字列直接パース・28桁除算・スケール整合・overflow） |
| `tests/spec/decimal-precision.test.ts` | Decimal の 15 シナリオ（基本演算・配列・関数引数・financial 計算） |

---

## 8. 既知の制限・将来課題

| 項目 | 現状 | 方針 |
|---|---|---|
| `Format()` の高精度対応 | Double 経由で 15 桁精度 | 実需要が出てから `formatNumber` を BigInt 対応に拡張 |
| `Dim d As Decimal` スカラーの自動強制 | 代入時に `VbaDecimal` に変換されない | `CDec()` を明示すれば格納可。型マップ拡張は他の型との整合性を要確認 |
| `JSON.stringify` との互換性 | `VbaCurrency` / `VbaDecimal` を含むオブジェクトは BigInt で TypeError | シリアライズ経路では `toString()` / `Number()` で変換してから渡す |
| Single との演算 | `Currency ⊕ Single` は Single の 2 進誤差を `fromNumber` 経由で吸収 | バンカーズ丸め統一。実 VBA との突き合わせテストで検証済み（`currency-bigint.test.ts`） |
