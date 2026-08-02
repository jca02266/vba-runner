---
type: analysis
title: 複数Finding評価の真因横断レビュー
description: 旧方式の複数Finding評価を、手順v1の真因仮説候補として分類した記録
tags:
  - evaluation
  - root-cause
  - audit
status: active
version: 1
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

現時点の判定は`hypothesis`であり、確定には可変arity、Null/Empty、関数固有範囲を壊さない最小設計検証が必要である。

2026-08-02の追加確認では、`BuiltinParamSpec.coerce`は`boolean`と`string`だけで、`bindCallArguments`は名前・個数・一部のホスト変換に限定されていた。したがって、次の設計境界を検証対象とする。

- 単純な必須/Optional Long引数は、対象型・丸め・Null/Empty方針をメタデータで表現できる。
- Compareの0/1、ChrB/RGBの範囲、CVErrのNull、InStr系の可変arityは、単純な共通coerceへ押し込まず、契約上の範囲または関数固有validatorとして残す。
- `bindOverloadedCallArguments`の引数個数による形式選択と、関数本体の評価順序を変更しない。

この境界が確認できるまでは、BuiltinParamSpec全体を置き換える実装タスクは開始しない。

### ファイルシステムのNull境界

対象は[EV-00337](../../evaluation/evaluations/EV-00337.md)と[EV-00338](../../evaluation/evaluations/EV-00338.md)である。

- CurDir/ChDirはNullをcwdまたはPath not foundへ変換していた。
- EOF/LOF/LOC/SeekとファイルI/O文は、ファイル番号のNullをホスト数値変換へ渡してError 94を失っていた。
- 両者は「ホスト変換より前にVBA Nullを検証し、Error 94へ写像する」不変条件で説明できる。

未監査のファイル番号・パス入口、Empty/Missing、名前付き引数、未オープン番号Error 52との境界を確認するまで`hypothesis`とする。

2026-08-02のソース監査では、文のファイル番号は`evaluateFileNumber`、組み込みのEOF/LOF/LOC/Seekは`toFileNumber`を通ることを確認した。FileAttrはNullを入口で検証してから`toVbaNumber`を使う。したがって、現時点で同じNull漏れの新しい入口は見つかっていないが、Empty/Missingとパス引数の全列挙は未完了である。

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

## 次のループ

1. BuiltinParamSpec候補について、引数契約を拡張する最小設計をソース上で検証する。
2. ファイルシステムNull境界について、未監査入口を機械列挙し、NullとEmpty/Missing/Error 52を分離する。
3. 各候補について、独立`challenge`または`confirm`レビューと未解決項目の解消を行う。
4. 真因が確定した候補だけ、新しい評価記録とcauseKeyを作成し、`evaluation/root-causes/RC-xxxxx.md`へ登録する。

旧EVを書き換えて原因キーを統合すること、`priorCauseKey`だけを根拠に真因を確定することは禁止する。
