---
type: Design Note
title: VBA 識別子分類の実装
description: Lexer と Parser が文脈に応じて VBA 識別子とキーワードを分類する実装方針。
tags: [vba, lexer, parser, identifiers, design]
status: stable
---

# VBA識別子分類の実装

## 目的

VBAのキーワードは、構文上の予約語であるものと、特定の構文位置では
キーワードとして働くが識別子としても使用できるものがある。本実装では、
Lexerのトークン化とParserの文脈依存の識別子判定を分離し、宣言構文ごとの
判定漏れを防いでいる。

## 仕様上の分類

MS-VBAL §3.3.5.2 の `reserved-identifier` は次のカテゴリから構成される。

```text
statement-keyword / marker-keyword / operator-identifier / special-form
reserved-type-identifier / reserved-name / literal-identifier / rem-keyword
reserved-for-implementation-use / future-reserved
```

カテゴリに属する代表例は次のとおりである。キーワードによっては複数の
カテゴリに属する（例: `New`, `Case`, `Write`）。

| カテゴリ | 代表例 |
|---|---|
| `statement-keyword` | `Call`, `Dim`, `For`, `Function`, `If`, `Set`, `Sub`, `Type`, `Write` |
| `marker-keyword` | `As`, `ByRef`, `ByVal`, `Each`, `New`, `Optional`, `Preserve`, `Then`, `To` |
| `operator-identifier` | `AddressOf`, `And`, `Is`, `Like`, `Mod`, `Not`, `Or`, `TypeOf`, `Xor` |
| `special-form` | `Array`, `Input`, `InputB`, `LBound`, `Scale`, `UBound` |
| `reserved-type-identifier` | `Boolean`, `Byte`, `Date`, `Double`, `Integer`, `Long`, `String`, `Variant` |
| `reserved-name` | `Abs`, `CDate`, `Debug`, `DoEvents`, `Len`, `Me`, `Sgn`, `String` |
| `literal-identifier` | `True`, `False`, `Nothing`, `Empty`, `Null` |
| `rem-keyword` | `Rem` |

完全なスペリング単位の対応表は、実装の
`VBA_KEYWORD_CATEGORIES`（`src/engine/lexer.ts`）で管理する。

## Lexerの実装

`VBA_KEYWORD_CATEGORIES` は、キーワードの小文字スペリングから仕様カテゴリ
配列への読み取り専用マップである。これはLexerのトークン化を置き換えるもの
ではなく、仕様分類をコード上で参照可能にするメタデータである。

```ts
VBA_KEYWORD_CATEGORIES.new
// ['marker-keyword', 'operator-identifier']
```

Lexerは従来どおり `KeywordStep` や `KeywordClass` などの専用トークンを生成する。
`KeywordStep` を `Identifier` に変更するような分類変更は行わない。

カテゴリのいずれにも属さない専用Lexerトークンは
`VBA_CONTEXTUAL_KEYWORDS` から取得できる。代表例は次のとおりである。

- `Output`, `Append`, `Random`, `Binary`
- `Access`, `Read`, `Text`, `Compare`, `Explicit`, `Base`, `Module`
- `Lib`, `Alias`, `PtrSafe`, `Step`
- `Mid`, `Kill`, `Width`, `Line`, `Reset`, `AppActivate`, `SendKeys`
- `Class`, `Collection`, `Error`, `Property`

## Parserの識別子判定

Parserの `isIdentifier(token)` は、通常の `Identifier` に加えて
`CONTEXTUAL_KW` に含まれるトークンを識別子として受け入れる。

この判定は次の宣言位置で共通利用される。

- `Dim` / `Const` の名前
- `Sub` / `Function` / `Property` の名前と引数
- `Enum` の名前とメンバー
- `Type` の名前とメンバー
- `For` / `For Each` の変数
- クラスメンバー

したがって、例えば次のコードは有効である。

```vba
Const STEP = 5
Dim Output As String
Function Property(ByVal Class As Long) As String
    Property = CStr(Class)
End Function
```

一方、`Function`、`Sub`、`Dim`、`Set` などの `statement-keyword` は、
通常の識別子位置では許可されない。

## `COMPAT_KW_EXPR` との違い

`Get`、`Put`、`Print`、`Seek`、`Lock`、`Unlock`、`Input` などは、仕様上は
予約語だが、既存VBAコードとの互換性のためメンバー名として許可する。

```vba
obj.Print value
worksheet.Get(index)
```

これは仕様準拠の contextual keyword ではないため、Parser内の
`COMPAT_KW_EXPR` で別管理する。`CONTEXTUAL_KW` や
`VBA_CONTEXTUAL_KEYWORDS` に追加して宣言名まで許可してはいけない。

## 型接尾辞との関係

Lexerはキーワード照合時に型接尾辞を除去するため、`Dim$` や `For%` は予約語
として扱われる。一方、contextual keywordの `Append$` などは識別子として
許可される。この挙動は `lexer-column.test.ts` などで検証する。

## 検証

`tests/spec/keyword-classification.test.ts` で次を検証する。

- 複数カテゴリに属するキーワードの分類
- 型名・リテラルの分類
- contextual keywordの分類外判定
- 既存の `Keyword*` トークン型の安定性
- `Shared` など予約語の予約状態

既存の contextual keyword 回帰テストと全体テストを合わせて、分類メタデータ
追加によるParser・Lexerの動作変更がないことを確認する。
