---
type: analysis
title: 複数Finding評価の真因横断レビュー
description: 旧方式の複数Finding評価を、手順v1の真因仮説候補として分類した記録
tags:
  - evaluation
  - root-cause
  - audit
status: completed
version: 2
verified: 2026-08-02
sources:
  - resource: ./ROOT_CAUSE_ANALYSIS.md
    description: 真因分析手順v1
---

# 複数Finding評価の真因横断レビュー

## 目的と扱い

この記録は、旧方式（`rootCauseProcedureVersion`がない、または`0`）の複数Finding評価を、手順v1で再分析するための分類表である。旧EVの`causeKey`や`priorCauseKey`は履歴として保持し、同じ名前や探索順の連鎖を新方式の真因確定とは扱わない。

現行の評価ストレージは、root causeの`originEvaluations`とEVの`causeKey`が完全一致する場合だけ登録を許可する。そのため、旧EVを改変して新しい原因へ昇格させず、以下の候補は独立レビュー後に新しい評価候補・RC記録として登録する。

## 共通経路候補

### BuiltinParamSpecと引数契約

対象は[EV-00302](../../evaluation/evaluations/EV-00302.md)〜[EV-00311](../../evaluation/evaluations/EV-00311.md)である。

- `src/engine/builtins.ts`の同じ組み込み登録層で、`ctx.toVbaNumber`、`ctx.round`、`Math.round`、`Math.trunc`、Null/範囲検証が関数ごとに手書きされている。
- `BuiltinParamSpec`と`bindCallArguments`は名前・個数と一部Boolean/String変換を扱うが、Long化、銀行丸め、Null/Empty/Missing、列挙値、範囲、可変arityの意味契約を表現しない。
- `normalizeCompare`でCompare検証だけは部分的に共通化されたが、Start/Limit/Count、必須引数の丸め、関数固有の範囲検証は分散している。

この候補は、次の部分仮説に分けて検証する。

1. Long化・銀行丸めの共通型変換責務がない。
2. Optionalの省略・Empty・Null・列挙値契約が宣言的でない。
3. BuiltinParamSpecがVBA引数契約を表現できず、上記二つを再発させる設計になっている。

この候補は、可変arity、Null/Empty、関数固有範囲を壊さない最小設計検証を完了するまで仮説として扱った。

2026-08-02の追加確認では、`BuiltinParamSpec.coerce`は`boolean`と`string`だけで、`bindCallArguments`は名前・個数・一部のホスト変換に限定されていた。したがって、次の設計境界を検証対象とする。

- 単純な必須/Optional Long引数は、対象型・丸め・Null/Empty方針をメタデータで表現できる。
- Compareの0/1、ChrB/RGBの範囲、CVErrのNull、InStr系の可変arityは、単純な共通coerceへ押し込まず、契約上の範囲または関数固有validatorとして残す。
- `bindOverloadedCallArguments`の引数個数による形式選択と、関数本体の評価順序を変更しない。

検証の結果、単純なLong変換だけをメタデータへ移し、Null/Missing保持、範囲、列挙値、可変arity、結果生成は各関数へ残す境界が成立した。この設計を[EV-00400](../../evaluation/evaluations/EV-00400.md)、[RC-00003](../../evaluation/root-causes/RC-00003.md)、[ROOT-00003](../../evaluation/remediations/ROOT-00003.md)へ記録した。

### 最小設計の検証マトリクス

| 代表関数 | 共通メタデータへ置ける責務 | 関数側に残す責務 | 確認する回帰 |
| --- | --- | --- | --- |
| `Left` / `Right` | 必須Long、銀行丸め、Null伝播方針 | 負値Error 5、文字列切り出し | 小数、Null、負値 |
| `Replace` | Optional Longの既定値、丸め、Null方針 | Compare 0/1、空検索、結果生成 | Start/Count小数、Empty、Null |
| `StrComp` | Optional Compareの丸めとMissing | Compare 0/1、Null文字列伝播 | Compare小数・無効値・Null |
| `InStr` / `InStrB` | 各形式のLong変換 | 引数個数による形式選択、Start下限、検索 | 2/3/4引数、名前付き、Null |
| `ChrB` / `RGB` / `CVErr` | 数値のVBA丸め | 範囲、clamp、Error番号固有規則 | 範囲外、Null、銀行丸め |

この表から、共通メタデータは「値の型変換」と「既定のNull/Empty方針」までに留め、
範囲・列挙値・可変arity・結果生成を自動化しないのが安全な境界と判断した。次の確認では、
この境界を表現する型を追加した場合に、既存の名前付き引数・評価順序・ByRef経路へ影響しないかを
最小の設計テストで確認する。

## 数値変換の追加ソース監査

`builtins.ts`の`Math.round`・`Math.floor`・`Math.trunc`と`ctx.toVbaNumber`の全使用箇所を確認した。Randomizeのseed生成、DateDiffの暦日・時刻差、DateValue/TimeValueのDate serial、文字列関数の位置計算などは、Long引数の共通変換とは異なるドメイン責務だった。今回の監査では仕様または実Excel結果なしに新しいバグとは判定せず、引数契約候補へ追加しない。

一方、文字列・数値組み込みの`ctx.round(ctx.toVbaNumber(...))`反復は、引数契約候補の根拠を補強する。次のループでは、代表4関数（Left、Replace、StrComp、InStr）の名前付き・位置引数テストを設計し、共通メタデータを導入しても評価順序と既存境界が保たれるかを確認する。

### ファイルシステムのNull境界

対象は[EV-00337](../../evaluation/evaluations/EV-00337.md)と[EV-00338](../../evaluation/evaluations/EV-00338.md)である。

- CurDir/ChDirはNullをcwdまたはPath not foundへ変換していた。
- EOF/LOF/LOC/SeekとファイルI/O文は、ファイル番号のNullをホスト数値変換へ渡してError 94を失っていた。
- 両者は「ホスト変換より前にVBA Nullを検証し、Error 94へ写像する」不変条件で説明できる。

未監査のファイル番号・パス入口、Empty/Missing、名前付き引数、未オープン番号Error 52との境界を確認した。

2026-08-02のソース監査では、文のファイル番号は`evaluateFileNumber`、組み込みのEOF/LOF/LOC/Seekは`toFileNumber`を通ることを確認した。FileAttrはNullを入口で検証してから`toVbaNumber`を使う。`filesystem-extra.test.ts`でNull、未オープン番号、パス境界を確認し、新しい共通原因候補は見つからなかった。

## 同一原因へ統合しない対象

次の評価はFindingが複数でも、今回の共通不変条件へは統合しない。

| 評価 | 分離理由 |
| --- | --- |
| EV-00193 | Select CaseのDecimal高精度比較という比較器の精度責務 |
| EV-00267 | Optional文字列変換と配列形状メタデータという異なる責務 |
| EV-00272 | TextStream投影プロパティとADODB.Stream lifecycleの別状態モデル |
| EV-00297 | Valの文字列走査・型サフィックス・基数符号拡張 |
| EV-00314 | Format/DatePart/DateDiffの週オプション契約 |
| EV-00339 | MkDirの既存パス・親階層というファイルシステム契約 |

## 完了判定

- BuiltinParamSpecの数値引数契約を代表関数で検証し、Long変換を共通化した。
- Null/Missing、関数固有範囲、可変arity、評価順序を回帰テストで維持した。
- EV-00302〜EV-00311の共通候補をRC-00003へ昇格し、ROOT-00003を完了した。
- EV-00337〜EV-00338は既存の共通Null検証とソース監査で新たな横展開漏れがないことを確認した。
- EV-00193、EV-00267、EV-00272、EV-00297、EV-00314、EV-00339は異なる責務として統合しない判定を完了した。
- `npm run typecheck --if-present`、対象回帰、`npm test`、`npm run eval -- validate`、`npm run check:docs`を実行した。

旧EVを書き換えて原因キーを統合すること、`priorCauseKey`だけを根拠に真因を確定することは禁止する。
