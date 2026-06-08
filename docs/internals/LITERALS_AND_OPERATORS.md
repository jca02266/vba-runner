# リテラル・型・演算子：仕様と実装

## 1. 演算子の優先順位

`parser.ts` の関数呼び出し階層がそのまま優先順位になる（下にある関数ほど結合が強い）。

| 優先度 | 演算子 | 種別 | パーサー関数 |
|--------|--------|------|-------------|
| 1（最低）| `Imp` | 論理包含 | `parseLogicalImp` |
| 2 | `Eqv` | 論理等価 | `parseLogicalEqv` |
| 3 | `Xor` | 排他的論理和 | `parseLogicalXor` |
| 4 | `Or` | 論理和 | `parseLogicalOr` |
| 5 | `And` | 論理積 | `parseLogicalAnd` |
| 6 | `Not`（単項） | 論理否定 | `parseLogicalNot` |
| 7 | `=` `<>` `Is` `Like` | 等値比較 | `parseEquality` |
| 8 | `<` `>` `<=` `>=` | 大小比較 | `parseRelational` |
| 9 | `&` | 文字列連結 | `parseConcatenation` |
| 10 | `+` `-` | 加減算 | `parseAdditive` |
| 11 | `Mod` | 剰余 | `parseModulo` |
| 12 | `\` | 整数除算 | `parseIntDivision` |
| 13 | `*` `/` | 乗除算 | `parseMultiplicative` |
| 14 | `-` `+`（単項） | 符号反転 | `parseUnary` |
| 15 | `^` | 冪乗 | `parseExponentiation` |
| 16（最高）| `.` `!` `()` `[]` | メンバー・呼び出し | `parsePrimary` |

### `^` の結合性

`parseExponentiation` は右辺に `parsePrimary` を呼ぶ（自己再帰しない）ため **左結合**。
VBA 仕様（MS-VBAL §5.6）も左結合と規定している。

```
2 ^ 2 ^ 3 = (2 ^ 2) ^ 3 = 4 ^ 3 = 64
```

---

## 2. `2^2` でスペースなしのときパースエラーになる理由

### 原因：レキサーが `^` を型サフィックスとして吸収する

VBA には数値リテラルの型宣言サフィックスとして `^`（LongLong）が存在する。
レキサーは数字を読み終えた直後に `%`, `&`, `@`, `!`, `#`, `^` が続いていると、
型サフィックスとしてリテラル文字列に取り込む（`lexer.ts` L.483-489）。

```
入力: 2^2
字句解析: Number("2^")   Number("2")   ← ^ が吸収され OperatorPower が出ない
入力: 2 ^ 2
字句解析: Number("2")  OperatorPower("^")  Number("2")  ← OK
```

| 書き方 | 結果 |
|--------|------|
| `2^2` | `2^` を Number として吸収 → パースエラー |
| `2 ^ 2` | `^` を OperatorPower として切り出し → OK（= 4） |
| `x^2` | 識別子の後は数値サフィックス文脈外 → OK |
| `2^x` | `2^` を Number として吸収、`x` が浮く → パースエラー |

### VBE では問題が起きない理由

VBE（Visual Basic Editor）は行確定時に演算子の前後へスペースを自動挿入する。
`2^2` と入力しても保存時点で `2 ^ 2` になるため、実務のコードベースでは起こらない。
VBE を経由しない手書き `.bas` やツール生成ファイルでは起こりうる。

---

## 3. 数値リテラルの型宣言サフィックス

### 一覧

| サフィックス | VBA 型 | 例 | 値域・精度 |
|-------------|--------|----|-----------|
| `%` | Integer | `100%` | 16 bit 整数（-32768〜32767） |
| `&` | Long | `100&` | 32 bit 整数（-2147483648〜2147483647） |
| `!` | Single | `1.5!` | 32 bit 浮動小数点（IEEE 754 single） |
| `#` | Double | `1.5#` | 64 bit 浮動小数点（IEEE 754 double） |
| `@` | Currency | `1.5@` | 固定小数点 4 桁（15.4 桁精度） |
| `^` | LongLong | `100^` | 64 bit 整数（64 bit 環境） |

### 識別子の型ヒントサフィックス（変数名）

変数名末尾にも型サフィックスを付けられる（`$` が追加、`^` は不可）。

| サフィックス | 型 |
|-------------|-----|
| `$` | String |
| `%` | Integer |
| `&` | Long |
| `!` | Single |
| `#` | Double |
| `@` | Currency |

例：`Dim name$` は `Dim name As String` と同義。

### 現在の実装の制限 ⚠️

パーサーはサフィックスを **除去するだけ** で、型情報を AST に保持しない（`parser.ts`）。

```ts
const cleanVal = token.value.replace(/[%&@!#^]$/, '');
expr = { type: 'NumberLiteral', value: Number(cleanVal) } as NumberLiteral;
// ↑ typeSuffix フィールドがないため型情報が消える
```

その結果、以下の型強制が行われない：

| 式 | VBA 仕様 | 現在の実装 |
|----|---------|-----------|
| `TypeName(100%)` | `Integer` | `Integer`（範囲一致で偶然正しい） |
| `TypeName(100&)` | `Long` | ✅ `Long`（修正済み） |
| `TypeName(1.5!)` | `Single` | ✅ `Single`（修正済み） |
| `TypeName(1.5@)` | `Currency` | ✅ `Currency`（修正済み） |
| `TypeName(100^)` | `LongLong` | ✅ `LongLong`（修正済み） |
| `100000%` | Error 6 Overflow | ✅ Error 6（修正済み） |

**実装済み**：`NumberLiteral` に `typeSuffix` / `isFloat` を追加し、`inferLiteralTypeName` / `inferLiteralVarType` がサフィックス優先で型を返すよう修正。

**値強制の状態**：

| サフィックス | TypeName/VarType | 値の強制 | 備考 |
|-------------|-----------------|---------|------|
| `%` Integer | ✅ 正確 | ✅ 銀行丸め + 範囲チェック | `32768%` → Error 6 |
| `&` Long    | ✅ 正確 | ✅ 銀行丸め + 範囲チェック | `2147483648&` → Error 6 |
| `!` Single  | ✅ 正確 | ⚠️ 値はそのまま（Math.fround なし） | VbaSingle ラッパー未実装のため |
| `#` Double  | ✅ 正確 | ✅ そのまま（no-op） | |
| `@` Currency| ✅ 正確 | ✅ 4桁丸め + 範囲チェック | ランタイムラッパーは未実装 |
| `^` LongLong| ✅ 正確 | ⚠️ 値はそのまま（BigInt 変換なし） | VbaLongLong ラッパー未実装のため |

---

## 4. サフィックスなしリテラルの型推定

### リテラル単体

VBA 仕様では数値リテラルの型はその値の範囲と小数点の有無で決まる。

| 式 | VBA 仕様 | 現在の実装 |
|----|---------|-----------|
| `TypeName(1)` | `Integer` | `Integer` ✓ |
| `TypeName(32767)` | `Integer` | `Integer` ✓ |
| `TypeName(32768)` | `Long` | `Long` ✓ |
| `TypeName(1.0)` | `Double` | `Integer` ✗ |
| `TypeName(1.5)` | `Double` | `Double` ✓ |

`1.0` の誤判定は JavaScript の `Number.isInteger(1.0) === true` による。
トークン文字列に `.` が含まれるかどうかで判定することで修正できる。

### 算術演算結果の型

VBA 仕様では演算結果の型はオペランドの型から決まる（型昇格）。

| 式 | VBA 仕様 | 現在の実装 |
|----|---------|-----------|
| `TypeName(1 + 1)` | `Integer` | `Double` ✗ |
| `TypeName(1 + 1.0)` | `Double` | `Double` ✓（偶然） |
| `TypeName(CInt(1) + CInt(1))` | `Integer` | `Double` ✗ |
| `TypeName(CLng(1) + CLng(1))` | `Long` | `Double` ✗ |
| `TypeName(CSng(1) + CSng(1))` | `Single` | `Double` ✗ |

**根本原因**：JavaScript の `number` は型情報を持たない。
算術演算の結果はすべて `typeof val === 'number'` となり、ランタイムの TypeName/VarType は
一律 `Double` を返す。

```ts
// evaluator.ts — 現状（ランタイム値の TypeName）
if (typeof val === 'number') return 'Double';   // TypeName
if (typeof val === 'number') return 5;           // VarType (vbDouble)
```

**修正するには**：`VbaInteger`・`VbaLong`・`VbaSingle` 等のラッパークラスを導入し、
演算ごとに VBA の型昇格ルール（MS-VBAL §6.1.2.2）を適用する必要がある。
全算術演算への影響が大きく、現時点では未実装。

### VBA の型昇格ルール（参考）

| 左辺 | 右辺 | `+` `-` `*` の結果 | `/` の結果 |
|------|------|-------------------|-----------|
| Integer | Integer | Integer（溢れたら Long → Double） | Double |
| Integer | Long | Long | Double |
| Long | Long | Long（溢れたら Double） | Double |
| Integer/Long | Single | Single | Single |
| Integer/Long/Single | Double | Double | Double |
| 任意 | Double | Double | Double |

---

## 5. Currency と Decimal の実装状況

### Currency（`@` / `CCur`）

| 項目 | 仕様 | 実装状況 |
|------|------|---------|
| 変換関数 `CCur(x)` | 4 小数桁に丸めて Currency 型を返す | ✅ 丸め実装済み（銀行丸め） |
| `TypeName(CCur(x))` | `"Currency"` | ✅ AST パスで `BUILTIN_RETURN_TYPES` から判定 |
| ランタイム型識別 | `TypeName(x)` で変数に格納後も `Currency` | ❌ 変数格納後は plain `number`、`TypeName` → `"Double"` |
| 算術型保持 | `CCur(1) + CCur(2)` の結果が Currency | ❌ 結果は plain `number` → `TypeName` → `"Double"` |
| 精度 | 固定小数点 15.4 桁 | ⚠️ IEEE 754 double のため厳密な固定小数点でない |
| ランタイムラッパークラス | 型情報を値に紐付けるクラス | ❌ 未実装（`CCur` は plain `number` を返す） |

**補足**：`TypeName(CCur(x))` が正しく `"Currency"` を返すのは、
TypeName が引数の AST を見て `CallExpression { callee: 'ccur' }` から
`BUILTIN_RETURN_TYPES` を参照しているため。
変数に代入した値はこの情報を持たない。

### Decimal（`CDec`）

| 項目 | 仕様 | 実装状況 |
|------|------|---------|
| 変換関数 `CDec(x)` | 28 桁固定精度の Decimal 型を返す | ✅ `VbaDecimal` ラッパーを返す |
| `TypeName(CDec(x))` | `"Decimal"` | ✅ `VbaDecimal` インスタンス検査で正しく判定 |
| ランタイム型識別 | 変数格納後も Decimal | ✅ `VbaDecimal` ラッパーが維持される |
| 精度 | 28 桁固定精度 | ❌ 内部は JavaScript `number`（IEEE 754）。`1.23456789012345678` → `1.2345678901234567`（17桁止まり） |
| 算術演算 | Decimal どうしの演算結果も Decimal | ❌ 算術演算で `VbaDecimal` がアンラップされ plain `number` になる |

**補足**：Currency と異なり `VbaDecimal` ラッパーは存在するため、
直接代入した変数の `TypeName` は正しく `"Decimal"` を返す。
ただし精度は IEEE 754 のまま。

---

## 6. `+` と `&` の違い

### `+` 演算子：左辺 × 右辺の型による挙動

| 左辺 \ 右辺 | 数値 | 数値文字列 (`"3"`) | 非数値文字列 (`"abc"`) | Null | Empty |
|------------|------|-------------------|----------------------|------|-------|
| **数値** | 数値加算<br>`1+2` → `3` | 右辺を数値変換して加算<br>`1+"3"` → `4` | **Type Mismatch**<br>`1+"abc"` → Error 13 | Null 伝播<br>`1+Null` → `Null` | 右辺を `0` 扱い<br>`1+Empty` → `1` |
| **数値文字列** | 左辺を数値変換して加算<br>`"3"+1` → `4` | **両辺とも String → 文字列連結**<br>`"3"+"4"` → `"34"` | 両辺とも String → 文字列連結<br>`"3"+"abc"` → `"3abc"` | Null 伝播 | 右辺を `""` 扱い<br>`"3"+Empty` → `"3"` |
| **非数値文字列** | **Type Mismatch**<br>`"abc"+1` → Error 13 | 両辺とも String → 文字列連結<br>`"abc"+"3"` → `"abc3"` | 文字列連結<br>`"abc"+"def"` → `"abcdef"` | Null 伝播 | 右辺を `""` 扱い<br>`"abc"+Empty` → `"abc"` |
| **Boolean** | 数値加算（True=-1）<br>`True+1` → `0` | **Type Mismatch** | **Type Mismatch** | Null 伝播 | `0` 扱い<br>`True+Empty` → `-1` |

**ルール**：両辺がともに String 型なら文字列連結。片方でも数値型なら数値加算を試みる。

### `&` 演算子：左辺 × 右辺の型による挙動

| 左辺 \ 右辺 | 数値 | 数値文字列 | 非数値文字列 | Null | Empty |
|------------|------|-----------|-------------|------|-------|
| **数値** | `"12"` | `"13"` | `"1abc"` | `"1"`（Null → `""`） | `"1"`（Empty → `""`） |
| **数値文字列** | `"31"` | `"34"` | `"3abc"` | `"3"` | `"3"` |
| **非数値文字列** | `"abc1"` | `"abc3"` | `"abcdef"` | `"abc"` | `"abc"` |
| **Boolean** | `"True1"` | `"True3"` | `"Trueabc"` | `"True"` | `"True"` |

**ルール**：常に両辺を `toDisplayString()` で文字列変換してから連結。Null・Empty は `""` 扱い（Null 伝播なし）。

### 判断フロー

```
+ 演算子:
  両辺がともに String 型?
  ├─ Yes → 文字列連結（数値文字列でも連結: "3"+"4" = "34"）
  └─ No  → 数値加算を試みる（変換不可文字列は Error 13）

& 演算子:
  常に両辺を文字列に変換して連結（Null / Empty → ""）
```

型が混在する場合は `&` を使うのが VBA の慣用。
`+` はオペランドの型によって加算か連結かが変わるため意図しない挙動の原因になりやすい。
