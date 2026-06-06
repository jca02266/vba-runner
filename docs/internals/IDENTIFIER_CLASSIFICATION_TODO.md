# 識別子分類の正規化 TODO

## 背景

`Const STEP = 5` などで contextual keyword を識別子として使えないバグが発生した（修正済）。
根本原因は `CONTEXTUAL_KW` チェックがパーサー各所に**散在・不統一**に適用されていること。
本 TODO は仕様に沿った識別子分類を定義し、実装を統一する作業を記録する。

---

## §3.3.5.2 の識別子分類（仕様）

```
reserved-identifier = statement-keyword
                    / marker-keyword
                    / operator-identifier
                    / special-form
                    / reserved-type-identifier
                    / reserved-name
                    / literal-identifier
                    / rem-keyword
                    / reserved-for-implementation-use
                    / future-reserved

IDENTIFIER = <lex-identifier> かつ NOT reserved-identifier
```

### 各カテゴリの具体的な値（仕様原文より抜粋）

| カテゴリ | キーワード |
|---|---|
| statement-keyword | Call, Case, Close, Const, Declare, Dim, Do, Else, ElseIf, End, Enum, Erase, Event, Exit, For, Friend, Function, Get, GoSub, GoTo, If, Implements, Input, Let, Lock, Loop, LSet, Next, On, Open, Option, Print, Private, Public, Put, RaiseEvent, ReDim, Resume, Return, RSet, Seek, Select, Set, Static, Stop, Sub, Type, Unlock, Wend, While, With, Write |
| marker-keyword | Any, As, ByRef, ByVal, Case, Each, Else, In, New, Shared, Until, WithEvents, Write, Optional, ParamArray, Preserve, Spc, Tab, Then, To |
| operator-identifier | AddressOf, And, Eqv, Imp, Is, Like, Mod, New, Not, Or, TypeOf, Xor |
| special-form | Array, Circle, Input, InputB, LBound, Scale, UBound |
| reserved-type-identifier | Boolean, Byte, Currency, Date, Double, Integer, Long, LongLong, LongPtr, Single, String, Variant |
| reserved-name | Abs, CBool, CByte, CCur, CDate, CDbl, CDec, CInt, CLng, CLngLng, CLngPtr, CSng, CStr, CVar, CVErr, Date, Debug, DoEvents, Fix, Int, Len, LenB, Me, PSet, Scale, Sgn, String |
| literal-identifier | True, False, Nothing, Empty, Null |
| rem-keyword | Rem |
| reserved-for-implementation-use | Attribute, LINEINPUT, VB_Base, VB_Control, VB_Creatable, ... |
| future-reserved | CDecl, Decimal, DefDec |

**IDENTIFIER として有効なキーワード**（上記のどのカテゴリにも属さない → `CONTEXTUAL_KW` に含めるべき）:

| トークン型 | キーワード | 備考 |
|---|---|---|
| `KeywordOutput` | Output | Open...For Output |
| `KeywordAppend` | Append | Open...For Append |
| `KeywordRandom` | Random | Open...For Random |
| `KeywordBinary` | Binary | Open...For Binary / Option Compare Binary |
| `KeywordAccess` | Access | Open...Access Read/Write |
| `KeywordRead` | Read | Open...Access Read |
| `KeywordText` | Text | Option Compare Text |
| `KeywordCompare` | Compare | Option Compare |
| `KeywordExplicit` | Explicit | Option Explicit |
| `KeywordBase` | Base | Option Base |
| `KeywordModule` | Module | Option Private Module |
| `KeywordLib` | Lib | Declare...Lib |
| `KeywordAlias` | Alias | Declare...Alias |
| `KeywordPtrSafe` | PtrSafe | Declare PtrSafe |
| `KeywordStep` | Step | For...Step |
| `KeywordMid` | Mid | Mid 関数 / Mid 代入文 |
| `KeywordKill` | Kill | Kill 文 |
| `KeywordWidth` | Width | Width# 文 |
| `KeywordLine` | Line | Line Input# 文 |
| `KeywordReset` | Reset | Reset 文 |
| `KeywordAppActivate` | AppActivate | AppActivate 文 |
| `KeywordSendKeys` | SendKeys | SendKeys 文 |
| `KeywordClass` | Class | Class 宣言の開始キーワード（spec の statement-keyword 一覧に不在���|
| `KeywordCollection` | Collection | 組み込みオブジェクト型名（reserved-identifier 一覧に不在）|
| `KeywordError` | Error | `On Error` 構文の構成要素（statement-keyword 一覧に不在）|
| `KeywordProperty` | Property | `Property Get/Set/Let` の開始キーワード（statement-keyword 一覧に不在）|

---

## 現状の問題点

### 1. `CONTEXTUAL_KW` チェックが散在・不統一

パーサー内でキーワードを識別子として受け入れる箇所が複数あり、各箇所で個別に `CONTEXTUAL_KW.has(token.type)` を書いている。新しい宣言構文を追加するたびにチェック漏れが起きやすい。

**確認済みのチェック漏れ（修正済）**:
- `parseConstDeclaration`: `Const STEP = 5` が失敗 → 修正済

**`CONTEXTUAL_KW` 未登録の追加判明キーワード**:
- `Class` / `Collection` / `Error` / `Property` — 仕様の `reserved-identifier` のどのカテゴリにも属さないが、Lexer が専用トークン型を持つため `CONTEXTUAL_KW` に入れないと識別子として使えない。Step 1 で要対応。

**未確認の可能性がある箇所**:
- `parseVarDeclaration`（Dim 宣言の変数名）
- `parseProcedureDeclaration`（Sub/Function 名）
- `parseEnumDeclaration`（Enum 名・メンバー名）
- `parseTypeDeclaration`（Type 名・メンバー名）
- `parseForStatement`（ループ変数名）
- `parseMemberExpression`（プロパティ名）

### 2. `COMPAT_KW_EXPR` との混在リスク

`Get/Put/Print/Seek/Lock/Unlock/Input` は仕様上 `statement-keyword`（= reserved-identifier）であり、本来は `IDENTIFIER` として使えない。ただし `obj.Print` や `ws.Get` のようなメンバーアクセスで実用上必要なため `COMPAT_KW_EXPR` として式コンテキストで許容している。

これは**仕様外の互換拡張**であり、`CONTEXTUAL_KW`（仕様準拠）と明確に分離して管理しなければならない。

---

## 実装 TODO

### ✅ Step 1: `CONTEXTUAL_KW` の網羅性検証

`Class / Collection / Error / Property` が未登録と判明。`CONTEXTUAL_KW_STRUCTURAL` グループとして追加した。

### ✅ Step 2: `isIdentifier()` ヘルパーメソッドを導入

```typescript
// parser.ts
private isIdentifier(token: Token): boolean {
    return token.type === TokenType.Identifier
        || Parser.CONTEXTUAL_KW.has(token.type);
}
```

### ✅ Step 3: 欠陥検出テストを先に書く

`tests/spec/contextual-keyword-as-identifier.test.ts` を作成済み。

### ✅ Step 4: 散在チェックを `isIdentifier()` で統一

- `parseConstDeclaration`: `isIdentifier()` に統一
- `parseVarDeclaration`: `isIdentifier()` に統一
- `parseProcedureDeclaration`: `isIdentifier()` に統一
- statement dispatcher: `Class`・`Error`・`Property` に `!= OperatorEquals` ガード追加
  （`Class = x` が Class宣言と誤判定されないようにする）

### ✅ Step 5: 全テスト通過を確認してコミット

全97テスト + 新規テスト通過。回帰なし。

### ✅ 残課題（完了）

- `parseEnumDeclaration`・`parseTypeDeclaration`・`parseForStatement` の変数名チェックも
  `isIdentifier()` に統一する → **commit `fd98345` で対応済み**
  - `parseTypeDeclaration`: `isIdentifier(nameToken)` に統一
  - `parseEnumDeclaration`: `isIdentifier(nameToken)` に統一
  - `parseForStatement`: `isIdentifier(idToken)` に統一

---

## 注意事項

- Lexer のトークン型は変更しない（`KeywordStep` → `Identifier` にしない）。Parser 側で吸収する。
- `COMPAT_KW_EXPR` は `CONTEXTUAL_KW` と混ぜない。仕様外拡張として明示的に分離したまま管理する。
- `KeywordShared` は marker-keyword（reserved-identifier）なので `CONTEXTUAL_KW` に入れてはいけない。
