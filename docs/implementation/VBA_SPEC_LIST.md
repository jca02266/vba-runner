---
type: Reference
title: VBA仕様実装一覧
description: MS-VBAL仕様に対するVBA構文、標準機能、ランタイム挙動の実装状況。
tags: [vba, specification, implementation, status]
status: stable
---

# VBA仕様実装一覧（MS-VBAL仕様書準拠）

この文書は、MS-VBAL仕様に定義されたVBAの構文・標準機能・ランタイム挙動に
関する現在の実装状況だけを管理する。過去に発見・修正した仕様バグの詳細は
[VBA_ENGINE_BUG_HISTORY.md](../bug-records/VBA_ENGINE_BUG_HISTORY.md) に分離している。
IDE機能、テスト支援、モック拡張、解析器の高度化など仕様準拠以外の開発課題は
[TODO.md](../todo/TODO.md) に記載する。

**進捗**: MS-VBAL 仕様書で個別セクションに名前が付いた **構文要素 / ステートメント / 式 / 標準ライブラリ関数 / 組み込みクラス** はすべて実装済みです。

ただし、これは「言語仕様書のリストに名前のある機能」の網羅率であって、**VBA のランタイム挙動（仕様書本文に書かれているが個別セクションを持たない暗黙の振る舞い）** までを 100% 満たしているわけではありません。実装が未確認のランタイム挙動については本ドキュメントの「[VBA ランタイム挙動](#vba-ランタイム挙動)」セクションを参照してください。

## 実装状況の更新方針

ランタイム挙動の検証を以下の手順で進めていきます。

1. **「VBA ランタイム挙動」セクションのチェックボックスを未着手の項目から順に実装**
   - 仕様書の関連箇所を確認した上で、Evaluator / Lexer / Parser の必要箇所を修正
2. **各項目の実装時に "VBA ベテランが書きそうな代表コード" をテストに必ず含める**
   - 単純な API 呼び出しの確認だけでなく、実務でハマりやすい組み合わせや暗黙の挙動を網羅
   - 例: Default Property を実装する際は、`Range("A1") = 10` のように `.Value` を省略するパターンを含める
3. **テスト完了後、本ファイルの該当チェックボックスを ☑ にする**
   - 制限事項があれば該当行に明記

---

仕様書: `spec/[MS-VBAL]-250520.docx`（テキスト版: `spec/MS-VBAL.txt`）
- **バージョン**: 2.4（2025-05-20）
- **URL**: https://learn.microsoft.com/en-us/openspecs/microsoft_general_purpose_programming_languages/ms-vbal/d5418146-0bd2-45eb-9c7a-fd9502722c74

凡例: ✅ 実装済み / ❌ 未実装 / ⚠️ 部分実装・恒久的制限 / 🔮 将来検討

---

## ✅ ARCH-1: パーサー 2 パス化 — プリスキャン廃止・BNF 準拠の構文的判別へ移行（実装済み）

**実装**:
- `collectUserProcNames()` / `isUserProcOverride()` プリスキャン機構を削除
- `Open` の判別: `hasFileOpenSyntaxAhead()`（行内に `For <mode>` があれば file I/O、なければ識別子ブランチ）
- `Print` / `Write` / `Input` / `Put` / `Get` / `Seek`: `peek(1) === #` のみで判別（BNF 上 marked-file-number が必須なため）
- `parseProcedureDeclaration`: `STATEMENT_KW_RESERVED`（Open/Close/Print/Input 等）をプロシージャ名として reject（§3.3.5.2 準拠）

- | `contextual-keyword-as-identifier.test.ts` (Section 6: statement-keyword をプロシージャ名に使うとコンパイルエラー)

---

## 言語実装の拡張機能

| 状態 | 優先度 | 機能 | 概要 | テスト |
|------|--------|------|------|--------|
| ✅ | P2 | 日本語識別子対応 | MS-VBAL §3.3.5 に従い `isAlpha()` を Unicode プロパティエスケープ (`\p{L}`) で拡張。パーサーの `isWordToken()` も同様に対応。`Dim 氏名 As String` 等の識別子・UDT・関数名で使用可能に | `unicode-identifiers.test.ts` |
| ⚠️ | P2 | VarPtr / StrPtr / ObjPtr | **恒久的制限**。構文と呼び出しは成立し、非ゼロのダミー `Long` を返すが、実メモリーアドレスは取得しない。`VarPtr` は呼び出しごとに異なる値を返し、`StrPtr` / `ObjPtr` も実体のバッファーアドレス・IUnknownポインターとの対応を保証しない。したがって `CopyMemory` / `RtlMoveMemory` 等のポインター演算、アドレス同一性の判定、Win32 API連携はできない。ネイティブメモリモデルを導入しない限り、この制限は解消しない。 | `varptr.test.ts` |

---

## 第5章：制御ステートメント (§5.4.2)

| 状態 | 優先度 | 機能 | 仕様書 | テスト |
|------|--------|------|--------|--------|
| ✅ | P0 | Call Statement | §5.4.2.1 | `call-named-params.test.ts` |
| ✅ | P1 | Call Statement — モジュール修飾呼び出し `Call Module.Proc` / `Call Module.Proc(args)` | §5.4.2.1 | `module-qualified-calls.test.ts` |
| ✅ | P1 | Parameter 型名のモジュール修飾 `ByVal arg As MSForms.ReturnInteger` | §5.3.1 | `module-qualified-type.test.ts` |
| ✅ | P1 | Named Parameters | §5.6.3.1.2 | `call-named-params.test.ts` |
| ✅ | P1 | Named Parameters — キーワードを引数名に使用（unrestricted-name, §5.6.13.1） | COM メソッド例: `.Add Type:=xlValidateList` | `call-named-params.test.ts` |
| ✅ | P0 | While Statement (`While...Wend`) | §5.4.2.2 | `while-wend.test.ts` |
| ✅ | P0 | For Statement (`For...Next`、Step対応) | §5.4.2.3 | `for-next.test.ts` |
| ✅ | P0 | For Each Statement (`For Each...In`) | §5.4.2.4 | `for-each.test.ts` |
| ✅ | P1 | Exit For Statement | §5.4.2.5 | `for-next.test.ts` |
| ✅ | P0 | Do Statement (`Do While/Until...Loop`, `Do...Loop While/Until`, `Do...Loop`) | §5.4.2.6 | `do-loop.test.ts` |
| ✅ | P1 | Exit Do Statement | §5.4.2.7 | `do-loop.test.ts` |
| ✅ | P0 | If Statement（複数行） | §5.4.2.8 | `if-then-else.test.ts` |
| ✅ | P0 | Single-line If Statement | §5.4.2.9 | `if-then-else.test.ts` |
| ✅ | P0 | Select Case Statement | §5.4.2.10 | `select-case.test.ts` |
| ✅ | P1 | Stop Statement | §5.4.2.11 | `stop-end.test.ts` |
| ✅ | P1 | GoTo Statement（汎用ラベルジャンプ） | §5.4.2.12 | `goto-statement.test.ts` |
| ✅ | P0 | On…GoTo Statement（複数ラベル選択） | §5.4.2.13 | `on-goto.test.ts` |
| ✅ | P2 | GoSub Statement | §5.4.2.14 | `on-goto.test.ts` |
| ✅ | P2 | Return Statement（GoSub用） | §5.4.2.15 | `on-goto.test.ts` |
| ✅ | P2 | On…GoSub Statement | §5.4.2.16 | `on-goto.test.ts` |
| ✅ | P1 | Exit Sub Statement | §5.4.2.17 | `exit-statements.test.ts` |
| ✅ | P1 | Exit Function Statement | §5.4.2.18 | `exit-statements.test.ts` |
| ✅ | P1 | Exit Property Statement | §5.4.2.19 | `exit-property.test.ts` |
| ✅ | P2 | RaiseEvent Statement | §5.4.2.20 | `raiseevent.test.ts` |
| ✅ | P0 | With Statement (`With...End With`) | §5.4.2.21 | `with-statement.test.ts` |
| ✅ | P1 | End Statement（プログラム終了） | §5.4.2.22 | `stop-end.test.ts` |
| ✅ | P1 | Assert Statement | §5.4.2.23 | `assert.test.ts`, `debug-assert.test.ts` |

## 第5章：データ操作ステートメント (§5.4.3)

| 状態 | 機能 | 仕様書 | テスト |
|------|------|--------|--------|
| ✅ | P0 | Local Variable Declarations (`Dim`) | §5.4.3.1 | `dim-const.test.ts` |
| ✅ | P0 | Local Constant Declarations (`Const`) | §5.4.3.2 | `dim-const.test.ts` |
| ✅ | P1 | ReDim Statement | §5.4.3.3 | `redim-erase.test.ts` |
| ✅ | P1 | Erase Statement | §5.4.3.4 | `redim-erase.test.ts` |
| ✅ | P1 | Mid Statement（代入形式: `Mid(str, n) = val`） | §5.4.3.5 | `mid-statement.test.ts` |
| ✅ | P2 | MidB / Mid$ / MidB$ Statement | §5.4.3.5 | `mid-variants.test.ts` |
| ✅ | P1 | LSet Statement | §5.4.3.6 | `lset-rset.test.ts` |
| ✅ | P1 | RSet Statement | §5.4.3.7 | `lset-rset.test.ts` |
| ✅ | P0 | Let Statement（代入） | §5.4.3.8 | `dim-const.test.ts` |
| ✅ | P0 | Set Statement（オブジェクト参照） | §5.4.3.9 | `dim-const.test.ts` |

## 第5章：エラーハンドリング (§5.4.4)

| 状態 | 優先度 | 機能 | 仕様書 | テスト |
|------|--------|------|--------|--------|
| ✅ | P0 | On Error Statement（GoTo / Resume Next / GoTo 0） | §5.4.4.1 | `error_handling.test.ts` |
| ✅ | P0 | Resume Statement | §5.4.4.2 | `error_handling.test.ts` |
| ✅ | P2 | Error Statement（エラー番号の発生） | §5.4.4.3 | `error-statement.test.ts` |

## 第5章：ファイル操作ステートメント (§5.4.5)

| 状態 | 優先度 | 機能 | 仕様書 | テスト |
|------|--------|------|--------|--------|
| ✅ | P1 | Open Statement | §5.4.5.1 | `filesystem.test.ts` |
| ✅ | P1 | Close / Reset Statements | §5.4.5.2 | `filesystem.test.ts` |
| ✅ | P2 | Seek Statement | §5.4.5.3 | `filesystem-extra.test.ts` |
| ✅ | P2 | Lock / Unlock Statements | §5.4.5.4/5 | (制限事項: スタブ化) | `fs-lock-width.test.ts` |
| ✅ | P2 | Line Input Statement | §5.4.5.6 | `filesystem.test.ts` |
| ✅ | P2 | Width Statement | §5.4.5.7 | (制限事項: スタブ化) | `fs-lock-width.test.ts` |
| ✅ | P2 | Print Statement（`Print #n`） | §5.4.5.8 | `filesystem.test.ts` |
| ✅ | P2 | Write Statement | §5.4.5.9 | `filesystem.test.ts` |
| ✅ | P2 | Input Statement | §5.4.5.10 | `filesystem.test.ts` |
| ✅ | P2 | Put Statement | §5.4.5.11 | 基本型・型付き配列・UDT配列を実装。物理レイアウトの実Excel照合は `EVAL_LOG.md` #100〜#102 参照 | `filesystem-extra.test.ts` |
| ✅ | P2 | Get Statement | §5.4.5.12 | 基本型・型付き配列・UDT配列を実装。物理レイアウトの実Excel照合は `EVAL_LOG.md` #100〜#102 参照 | `filesystem-extra.test.ts` |

## 第5章：宣言 (§5.2, §5.3)

| 状態 | 優先度 | 機能 | 仕様書 | テスト |
|------|--------|------|--------|--------|
| ✅ | P0 | Sub / Function 宣言 | §5.3.1.6 | `procedure-declarations.test.ts` |
| ✅ | P0 | Parameter Lists（ByVal, ByRef, Optional, ParamArray） | §5.3.1.5 | `parameter-lists.test.ts` |
| ✅ | P0 | Property Get / Let / Set 宣言 | §5.3.1.7 | `property.test.ts` |
| ✅ | P2 | Event Handler 宣言 | §5.3.1.8 | `raiseevent.test.ts` |
| ✅ | P0 | Variable Declarations（`Dim`） | §5.2.3.1 | `dim-const.test.ts` |
| ✅ | P1 | Fixed-Length String Declaration（`Dim s As String * N`） | §5.2.3.1.4 | `fixed-length-string.test.ts` （初期値 NUL×N・代入時スペースパディング・切り捨て・UDT メンバー対応）|
| ✅ | P1 | Public / Private / Friend スコープ制御（パースのみ、アクセス制御なし） | §5.2.3.1 / §5.3.1.1 | `scope-private.test.ts` |
| ✅ | P1 | Static キーワード（変数・プロシージャ） | §5.3.1.2 | `static.test.ts` |
| ✅ | P0 | Const Declarations | §5.2.3.2 | `dim-const.test.ts` |
| ✅ | P1 | User Defined Type Declarations（`Type`） | §5.2.3.3 | `udt-type.test.ts` （`Width`/`Height`/`Left`/`Top`/`Name`/`Value` など予約語メンバー名も対応済み） |
| ✅ | P1 | Enum Declarations | §5.2.3.4 | `enum.test.ts` |
| ✅ | P0 | External Procedure Declaration（`Declare`） | §5.2.3.5 | (制限事項: JSランタイム上ではスタブとして登録) | `declare.test.ts` |
| ✅ | P0 | Class Module（OOP） | §5.2.4 | `class-module.test.ts` |
| ✅ | P1 | Implements Directive | §5.2.4.2 | `implements.test.ts` |
| ✅ | P2 | Event Declaration | §5.2.4.3 | `raiseevent.test.ts` |
| ✅ | P1 | Option Explicit（静的解析 + 実行時エラー） | §5.2 | `option-explicit.test.ts` |
| ✅ | P0 | Option Compare | §5.2.1.1 | `option_compare.test.ts` |
| ✅ | P1 | Option Base | §5.2.1.2 | `option_base.test.ts` |
| ✅ | P1 | Option Private Module | §5.2.1.4 | `option-private.test.ts` |
| ✅ | P0 | Attributes (VB_Name, etc.) | §5.2.3.1.6 / §5.2.4.1 | (制限事項: パースのみ。実行時は無視) | `ui_attr.test.ts` |
| ✅ | P0 | Date Literals (#mm/dd/yyyy#) | §3.3.3.3 | `number_literals.test.ts` |
| ✅ | P2 | FOREIGN-NAME `[identifier]` 構文 — 予約語をプロシージャ呼び出しに使用（例: `[End]()`）。定義側（`Sub [End]`）は実 VBA でも不可。未定義でも Option Explicit エラーにならない挙動を含む | §3.3.5.2 | `foreign-name.test.ts` |

## 第5章：演算子 (§5.6.9)

| 状態 | 優先度 | 機能 | 仕様書 | テスト |
|------|--------|------|--------|--------|
| ✅ | P1 | 算術演算子（`+`, `-`, `*`, `/`, `\`, `Mod`, `^`） | §5.6.9.3 | `math-module.test.ts` |
| ✅ | P1 | 文字列連結（`&`） | §5.6.9.4 | `operators-extra.test.ts` |
| ✅ | P1 | 比較演算子（`=`, `<>`, `<`, `>`, `<=`, `>=`） | §5.6.9.5 | `operators-extra.test.ts` |
| ✅ | P1 | Like 演算子 | §5.6.9.6 | (正規表現エスケープおよび[!charlist]対応) | `operators-extra.test.ts` |
| ✅ | P1 | Is 演算子（オブジェクト参照比較） | §5.6.9.7 | `typeof-is.test.ts` |
| ✅ | P1 | Not 演算子 | §5.6.9.8.1 | (制限事項: Boolean演算時は専用ラッパーを返し、JS側での直接評価には注意が必要) | `logical-operators.test.ts` |
| ✅ | P1 | And 演算子 | §5.6.9.8.2 | (制限事項: 同上) | `logical-operators.test.ts` |
| ✅ | P1 | Or 演算子 | §5.6.9.8.3 | (制限事項: 同上) | `logical-operators.test.ts` |
| ✅ | P1 | Xor 演算子 | §5.6.9.8.4 | (制限事項: 同上) | `logical-operators.test.ts` |
| ✅ | P1 | Eqv 演算子 | §5.6.9.8.5 | (制限事項: 同上) | `logical-operators.test.ts` |
| ✅ | P1 | Imp 演算子 | §5.6.9.8.6 | (制限事項: 同上) | `logical-operators.test.ts` |
| ✅ | P1 | TypeOf…Is 式 | §5.6.7 | `typeof-is.test.ts` |
| ✅ | P1 | Dictionary Access 式（`!` 演算子） | §5.6.14 | `dictionary-bang.test.ts` |
| ✅ | P2 | AddressOf 式 | §5.6.16.8 | `addressof.test.ts` |

## 第6章：標準ライブラリ — 変換関数 Conversion Module (§6.1.2.3)

| 状態 | 優先度 | 関数 | 仕様書 | テスト |
|------|--------|------|--------|--------|
| ✅ | P1 | CBool | §6.1.2.3.1.1 | (制限事項: Boolean専用ラッパーを返し、JS側での直接評価には注意が必要) | `conversion-extra.test.ts` |
| ✅ | P1 | CByte | §6.1.2.3.1.2 | `cbyte.test.ts` |
| ✅ | P2 | CCur | §6.1.2.3.1.3 | `ccur.test.ts` |
| ✅ | P0 | CDate / CVDate | §6.1.2.3.1.4 | `cdate.test.ts`, `conversion_date.test.ts` |
| ✅ | P1 | CDbl | §6.1.2.3.1.5 | `conversion-extra.test.ts` |
| ✅ | P2 | CDec | §6.1.2.3.1.6 | `cdec.test.ts` |
| ✅ | P1 | CInt | §6.1.2.3.1.7 | `conversion-extra.test.ts` |
| ✅ | P1 | CLng | §6.1.2.3.1.8 | `conversion-extra.test.ts` |
| ✅ | P2 | CLngLng | §6.1.2.3.1.9 | `clnglng.test.ts` |
| ✅ | P2 | CLngPtr | §6.1.2.3.1.10 | 64bit 環境で CLngLng と同等 | `clnglng.test.ts` |
| ✅ | P2 | CSng | §6.1.2.3.1.11 | `csng.test.ts` |
| ✅ | P1 | CStr | §6.1.2.3.1.12 | `conversion-extra.test.ts` |
| ✅ | P2 | CVar | §6.1.2.3.1.13 | `conversion-extra.test.ts` |
| ✅ | P2 | CVErr | §6.1.2.3.1.14 | `cverr.test.ts` |
| ✅ | P2 | Error / Error$ | §6.1.2.3.1.15 | `error_function.test.ts` |
| ✅ | P1 | Fix | §6.1.2.3.1.16 | `str_fix.test.ts` |
| ✅ | P1 | Hex / Hex$ | §6.1.2.3.1.17 | `number_literals.test.ts` |
| ✅ | P1 | Int | §6.1.2.3.1.18 | `conversion-int-val.test.ts` |
| ✅ | P1 | Oct / Oct$ | §6.1.2.3.1.19 | `number_literals.test.ts` |
| ✅ | P1 | Str / Str$ | §6.1.2.3.1.20 | `str_fix.test.ts` |
| ✅ | P1 | Val | §6.1.2.3.1.21 | `conversion-int-val.test.ts` |
| ✅ | P1 | TypeName | §6.1.2.3.1.22 | 宣言型変数・組み込み変換関数(CLng等)・UBound/Len等の固定戻り型関数・ユーザー定義関数の宣言戻り型を正しく返す。Variant変数は代入時に RHS の AST からサブタイプを追跡（`v=42`→Integer、`v=40000+1`→Long、`v=6/2`→Double、`v=6\2`→Integer）し、`TypeName(v)` / 直接式 `TypeName(1+1)` の両形式で正しく返す | `typename.test.ts`, `function-return-coercion.test.ts` |

## 第6章：標準ライブラリ — 日付/時刻 DateTime Module (§6.1.2.4)

| 状態 | 優先度 | 関数/プロパティ | 仕様書 | テスト |
|------|--------|----------------|--------|--------|
| ✅ | P1 | DateAdd | §6.1.2.4.1.1 | `datetime.test.ts` |
| ✅ | P1 | DateDiff | §6.1.2.4.1.2 | `datetime.test.ts` |
| ✅ | P1 | DatePart | §6.1.2.4.1.3 | `datetime.test.ts` |
| ✅ | P0 | DateSerial | §6.1.2.4.1.4 | `datetime.test.ts` |
| ✅ | P1 | DateValue | §6.1.2.4.1.5 | `datetime.test.ts` |
| ✅ | P1 | Day | §6.1.2.4.1.6 | `datetime.test.ts` |
| ✅ | P1 | Hour | §6.1.2.4.1.7 | `datetime.test.ts` |
| ✅ | P1 | Minute | §6.1.2.4.1.8 | `datetime.test.ts` |
| ✅ | P1 | Month | §6.1.2.4.1.9 | `datetime.test.ts` |
| ✅ | P1 | Second | §6.1.2.4.1.10 | `datetime.test.ts` |
| ✅ | P1 | TimeSerial | §6.1.2.4.1.11 | `datetime.test.ts` |
| ✅ | P1 | TimeValue | §6.1.2.4.1.12 | `datetime.test.ts` |
| ✅ | P1 | Weekday | §6.1.2.4.1.13 | `datetime.test.ts` |
| ✅ | P1 | Year | §6.1.2.4.1.14 | `datetime.test.ts` |
| ✅ | P1 | Date / Date$ | §6.1.2.4.2.2 | `datetime.test.ts` |
| ✅ | P1 | Now | §6.1.2.4.2.3 | `datetime.test.ts` |
| ✅ | P1 | Time / Time$ | §6.1.2.4.2.4 | `datetime.test.ts` |
| ✅ | P1 | Timer | §6.1.2.4.2.5 | `datetime.test.ts` |
| ✅ | P2 | WeekdayName | §6.1.2.4.1.15 | `datetime_names.test.ts` |
| ✅ | P2 | MonthName | §6.1.2.4.1.16 | `datetime_names.test.ts` |

## 第6章：標準ライブラリ — 数学関数 Math Module (§6.1.2.10)

| 状態 | 優先度 | 関数 | 仕様書 | テスト |
|------|--------|------|--------|--------|
| ✅ | P1 | Abs | §6.1.2.10.1.1 | `math-module.test.ts` |
| ✅ | P2 | Atn | §6.1.2.10.1.2 | `math-module.test.ts` |
| ✅ | P1 | Cos | §6.1.2.10.1.3 | `math-module.test.ts` |
| ✅ | P2 | Exp | §6.1.2.10.1.4 | `math-module.test.ts` |
| ✅ | P2 | Log | §6.1.2.10.1.5 | `math-module.test.ts` |
| ✅ | P1 | Rnd | §6.1.2.10.1.6 | `math-module.test.ts` |
| ✅ | P1 | Round | §6.1.2.10.1.7 | `math-module.test.ts` |
| ✅ | P2 | Sgn | §6.1.2.10.1.8 | `math-module.test.ts` |
| ✅ | P1 | Sin | §6.1.2.10.1.9 | `math-module.test.ts` |
| ✅ | P1 | Sqr | §6.1.2.10.1.10 | `math-module.test.ts` |
| ✅ | P1 | Tan | §6.1.2.10.1.11 | `math-module.test.ts` |
| ✅ | P1 | Randomize | §6.1.2.10.2.1 | `math-module.test.ts` |

## 第6章：標準ライブラリ — 文字列関数 Strings Module (§6.1.2.11)

| 状態 | 優先度 | 関数 | 仕様書 | テスト |
|------|--------|------|--------|--------|
| ✅ | P1 | Asc / AscW | §6.1.2.11.1.1 | `builtin-strings.test.ts` |
| ✅ | P1 | Chr / Chr$ / ChrW | §6.1.2.11.1.4/6 | `builtin-strings.test.ts` |
| ✅ | P1 | Filter | §6.1.2.11.1.7 | `filter.test.ts` (Bug CL: 多次元配列→error 13) |
| ✅ | P0 | Format / Format$ | §6.1.2.11.1.8/9 | `format.test.ts` |
| ✅ | P1 | FormatCurrency / FormatNumber / FormatPercent / FormatDateTime | — | `builtins.test.ts` |
| ✅ | P0 | InStr / InStrB | §6.1.2.11.1.14 | `builtin-strings.test.ts` (Bug BH: typeof判定誤分類修正、Null Start/Compare→error 94) |
| ✅ | P1 | InStrRev | §6.1.2.11.1.15 | `instrrev.test.ts` (Bug CJ: Null Start/Compare→error 94、Start<1→error 5、空StringCheck/Start>Len優先→0) |
| ✅ | P1 | Join | §6.1.2.11.1.16 | `split-join.test.ts` |
| ✅ | P1 | LCase / LCase$ | §6.1.2.11.1.17/18 | `builtin-strings.test.ts` |
| ✅ | P1 | Left / Left$ | §6.1.2.11.1.19/20 | `builtin-strings.test.ts` |
| ✅ | P0 | Len | §6.1.2.11.1.22 | `builtin-strings.test.ts` |
| ✅ | P1 | LTrim / RTrim | §6.1.2.11.1.23 | `trim.test.ts` |
| ✅ | P1 | Trim / Trim$ | §6.1.2.11.1.23 | `trim.test.ts` |
| ✅ | P0 | Mid / Mid$ | §6.1.2.11.1.25/26 | `builtin-strings.test.ts` |
| ✅ | P1 | Replace | §6.1.2.11.1.29 | `builtins.test.ts`, `builtin-strings.test.ts` (Bug BG: Expression=Null→error 94、Start<1→error 5) |
| ✅ | P1 | Right / Right$ | §6.1.2.11.1.30/31 | `builtin-strings.test.ts` |
| ✅ | P1 | Space / Space$ | §6.1.2.11.1.33/34 | `builtin-strings.test.ts` |
| ✅ | P1 | Split | §6.1.2.11.1.35 | `split-join.test.ts` (Bug CK: 空Expression→空配列、空Delimiter→1要素、Compare実装、Null引数→error 94) |
| ✅ | P1 | StrComp | §6.1.2.11.1.36 | `strcomp.test.ts` |
| ✅ | P2 | StrConv | §6.1.2.11.1.37 | `strconv.test.ts`, `strconv_null.test.ts` |
| ✅ | P1 | LeftB / RightB / MidB | §6.1.2.11.1.x | `mid-variants.test.ts` |
| ✅ | P1 | String / String$ | §6.1.2.11.1.38/39 | `builtin-strings.test.ts` |
| ✅ | P1 | StrReverse | §6.1.2.11.1.40 | `strreverse.test.ts` |
| ✅ | P1 | UCase / UCase$ | §6.1.2.11.1.41/42 | `builtin-strings.test.ts` |

## 第6章：標準ライブラリ — 情報関数 Information Module (§6.1.2.7)

| 状態 | 優先度 | 関数 | 仕様書 | テスト |
|------|--------|------|--------|--------|
| ✅ | P1 | IsArray | §6.1.2.7.1.2 | `info_functions.test.ts` |
| ✅ | P0 | IsDate | §6.1.2.7.1.3 | `info_functions.test.ts` |
| ✅ | P1 | IsEmpty | §6.1.2.7.1.4 | `info_functions.test.ts` |
| ✅ | P1 | IsError | §6.1.2.7.1.5 | `info_functions.test.ts` |
| ✅ | P1 | IsMissing | §6.1.2.7.1.6 | `info_functions.test.ts` |
| ✅ | P1 | On Error | §5.4.4.1 | `error_handling.test.ts` |
| ✅ | P1 | Resume | §5.4.4.2 | `error_handling.test.ts` |
| ✅ | P1 | Err Object | §6.1.3.2 | `error_handling.test.ts` |
| ✅ | P1 | IsNull | §6.1.2.7.1.7 | `info_functions.test.ts` |
| ✅ | P1 | IsNumeric | §6.1.2.7.1.8 | `info_functions.test.ts` |
| ✅ | P1 | IsObject | §6.1.2.7.1.9 | `info_functions.test.ts` |
| ✅ | P0 | TypeName | §6.1.2.7.1.12 | UDT・Class対応。Variant変数の動的サブタイプ追跡（代入時 AST 解析＋BinaryExpression型昇格規則）対応 | `typename.test.ts` |
| ✅ | P1 | VarType | §6.1.2.7.1.13 | Variant変数の動的サブタイプ追跡対応 | `vartype.test.ts` |
| ✅ | P0 | Err オブジェクト | §6.1.2.4 | (Number, Source, Description等、Raise/Clearメソッド対応) | `error_handling.test.ts` |

## 第6章：標準ライブラリ — インタラクション Interaction Module (§6.1.2.8)

| 状態 | 優先度 | 関数 | 仕様書 | テスト |
|------|--------|------|--------|--------|
| ✅ | P1 | Choose | §6.1.2.8.1.2 | `choose_switch.test.ts` |
| ✅ | P0 | CreateObject (拡張) | §6.1.2.8.1.4 | (制限事項: Scripting.Dictionary等の主要オブジェクトのスタブ・フェイク実装のみ対応) | `createobject.test.ts` |
| ✅ | P2 | Environ / Environ$ | §6.1.2.8.1.6 | `environ.test.ts` |
| ✅ | P1 | GetObject | §6.1.2.8.1.9 | `getobject.test.ts` |
| ✅ | P1 | IIf | §6.1.2.8.1.11 | `extra_fns.test.ts` |
| ✅ | P0 | InputBox | §6.1.2.8.1.12 | (制限事項: スタブ化。コンソール出力と固定値返却) | `interaction.test.ts` |
| ✅ | P0 | MsgBox | §6.1.2.8.1.13 | (制限事項: スタブ化。コンソール出力と固定値返却) | `interaction.test.ts` |
| ✅ | P1 | Shell | §6.1.2.8.1.15 | (制限事項: スタブ化。コマンド実行はせずログ出力のみ) | `shell.test.ts` |
| ✅ | P1 | Switch | §6.1.2.8.1.16 | `choose_switch.test.ts` |
| ✅ | P2 | GetSetting / SaveSetting | §6.1.2.8.1.10 / .14 | (制限事項: 仮想レジストリ) | `registry.test.ts` |
| ✅ | P2 | DeleteSetting / GetAllSettings | §6.1.2.8.1.5 / .8 | (制限事項: 仮想レジストリ) | `registry.test.ts` |
| ✅ | P1 | DoEvents | §6.1.2.8.1.5 | (制限事項: スタブ化) | `extra_fns.test.ts` |
| ✅ | P1 | AppActivate | §6.1.2.8.1.1 | (制限事項: スタブ化) | `extra_fns.test.ts` |
| ✅ | P1 | SendKeys | §6.1.2.8.1.14 | (制限事項: スタブ化) | `extra_fns.test.ts` |

## 第6章：標準ライブラリ — 配列関連

| 状態 | 優先度 | 関数 | 仕様書 | テスト |
|------|--------|------|--------|--------|
| ✅ | P0 | UBound | §6（標準ライブラリ） | `array-functions.test.ts` |
| ✅ | P0 | LBound | §6（標準ライブラリ） | `array-functions.test.ts` |
| ✅ | P0 | Array関数 | §6（標準ライブラリ） | `array-functions.test.ts` |

## 第6章：標準ライブラリ — ファイルシステム FileSystem Module (§6.1.2.5)

| 状態 | 優先度 | 関数/サブプロシージャ | 仕様書 | テスト |
|------|--------|----------------|--------|--------|
| ✅ | P2 | CurDir / CurDir$ | §6.1.2.5.1.1 | `curdir.test.ts` |
| ✅ | P2 | Dir | §6.1.2.5.1.2 | `mkdir_rmdir.test.ts` |
| ✅ | P2 | EOF | §6.1.2.5.1.3 | `filesystem.test.ts` |
| ✅ | P2 | FileAttr | §6.1.2.5.1.4 | (制限事項: スタブ化) | `filesystem-extra.test.ts` |
| ✅ | P2 | FileDateTime | §6.1.2.5.1.5 | `filesystem-extra.test.ts` |
| ✅ | P2 | FileLen | §6.1.2.5.1.6 | `filesystem-extra.test.ts` |
| ✅ | P1 | FreeFile | §6.1.2.5.1.7 | `freefile.test.ts` |
| ✅ | P2 | Loc | §6.1.2.5.1.8 | `filesystem.test.ts` |
| ✅ | P2 | LOF | §6.1.2.5.1.9 | `filesystem.test.ts` |
| ✅ | P2 | Seek | §6.1.2.5.1.10 | `filesystem-extra.test.ts` |
| ✅ | P2 | ChDir | §6.1.2.5.2.1 | `mkdir_rmdir.test.ts` |
| ✅ | P2 | ChDrive | §6.1.2.5.2.2 | (制限事項: スタブ化) | `mkdir_rmdir.test.ts` |
| ✅ | P2 | FileCopy | §6.1.2.5.2.3 | `mkdir_rmdir.test.ts` |
| ✅ | P2 | Kill | §6.1.2.5.2.4 | `mkdir_rmdir.test.ts` |
| ✅ | P2 | MkDir | §6.1.2.5.2.5 | `mkdir_rmdir.test.ts` |
| ✅ | P2 | RmDir | §6.1.2.5.2.6 | `mkdir_rmdir.test.ts` |
| ✅ | P2 | SetAttr | §6.1.2.5.2.7 | (制限事項: スタブ化) | `mkdir_rmdir.test.ts` |

## 第6章：標準ライブラリ — 財務 Financial Module (§6.1.2.6)

| 状態 | 優先度 | 関数 | 仕様書 | テスト |
|------|--------|------|--------|--------|
| ✅ | P2 | DDB | §6.1.2.6.1.1 | `financial.test.ts` |
| ✅ | P2 | FV | §6.1.2.6.1.2 | `financial.test.ts` |
| ✅ | P2 | IPmt | §6.1.2.6.1.3 | `financial.test.ts` |
| ✅ | P2 | IRR | §6.1.2.6.1.4 | `financial.test.ts` |
| ✅ | P2 | MIRR | §6.1.2.6.1.5 | `financial.test.ts` |
| ✅ | P2 | NPer | §6.1.2.6.1.6 | `financial.test.ts` |
| ✅ | P2 | NPV | §6.1.2.6.1.7 | `financial.test.ts` |
| ✅ | P2 | Pmt | §6.1.2.6.1.8 | `financial.test.ts` |
| ✅ | P2 | PPmt | §6.1.2.6.1.9 | `financial.test.ts` |
| ✅ | P2 | PV | §6.1.2.6.1.10 | `financial.test.ts` |
| ✅ | P2 | Rate | §6.1.2.6.1.11 | `financial.test.ts` |
| ✅ | P2 | SLN | §6.1.2.6.1.12 | `financial.test.ts` |
| ✅ | P2 | SYD | §6.1.2.6.1.13 | `financial.test.ts` |

## 第6章：組み込みクラス (§6.1.3)

### Collection Object (§6.1.3.1)

| 状態 | 優先度 | 機能 | 仕様書 | テスト |
|------|--------|------|--------|--------|
| ✅ | P1 | Count | §6.1.3.1.1.1 | `collection-remove.test.ts` |
| ✅ | P1 | Item | §6.1.3.1.1.2 | `collection-remove.test.ts` |
| ✅ | P1 | Add | §6.1.3.1.2.1 | `collection-remove.test.ts` |
| ✅ | P1 | Remove | §6.1.3.1.2.2 | `collection-remove.test.ts` |

### Err Class (§6.1.3.2)

| 状態 | 優先度 | 機能 | 仕様書 | テスト |
|------|--------|------|--------|--------|
| ✅ | P1 | Clear | §6.1.3.2.1.1 | `err-object-full.test.ts` |
| ✅ | P1 | Raise | §6.1.3.2.1.2 | `err-object-full.test.ts` |
| ✅ | P1 | Description | §6.1.3.2.2.1 | `err-object-full.test.ts` |
| ✅ | P2 | HelpContext | §6.1.3.2.2.2 | `err-object-full.test.ts` |
| ✅ | P2 | HelpFile | §6.1.3.2.2.3 | `err-object-full.test.ts` |
| ✅ | P2 | LastDllError | §6.1.3.2.2.4 | `err-object-full.test.ts` |
| ✅ | P1 | Number | §6.1.3.2.2.5 | `err-object-full.test.ts` |
| ✅ | P1 | Source | §6.1.3.2.2.6 | `err-object-full.test.ts` |

---

## 実装優先度ガイド（履歴）

MS-VBAL仕様書に個別セクションを持つ機能の実装優先度は、初期開発時の履歴として管理していた。
現在の実装状況は各仕様セクションの状態表を正とし、未実装の製品機能ロードマップは `docs/todo/TODO.md` を参照する。


## 仮想ファイルシステム (VFS) の実装状況と課題

過去の仕様バグの詳細な症状・原因・修正履歴は [VBA_ENGINE_BUG_HISTORY.md](../bug-records/VBA_ENGINE_BUG_HISTORY.md) に記録する。

Webブラウザおよびテスト環境向けの仮想ファイルシステム (`MemoryFileSystem`) の実装状況です。

### 機能マトリックス

| 状態 | 機能カテゴリ | 項目 | 説明 |
| :---: | :--- | :--- | :--- |
| ✅ | **基本 I/O** | `readSync` / `writeSync` | バッファー/文字列の同期読み書き。ハンドル位置の自動更新対応。 |
| ✅ | | `openSync` / `closeSync` | ファイルオープン。`w` フラグ時の truncation (切り詰め)、`a` フラグ時の末尾追記に対応済み。修正履歴は [VBA_ENGINE_BUG_HISTORY.md](../bug-records/VBA_ENGINE_BUG_HISTORY.md) を参照。 |
| ✅ | **ファイル管理** | `unlinkSync` | `Kill` ステートメントの基盤。現状は完全一致のみ。 |
| ✅ | | `copyFileSync` | `FileCopy` ステートメントの基盤。 |
| ✅ | **ディレクトリ** | `mkdirSync` | ディレクトリ作成。`recursive: true` 対応。 |
| ✅ | | `readdirSync` | ディレクトリ一覧取得。`Dir` 関数の基盤。 |
| ✅ | | `rmdirSync` | ディレクトリ削除。 |
| ✅ | **メタデータ** | `statSync` | `size`, `mtime`, `birthtime`, `mode` の取得。 |
| ✅ | | `existsSync` | ファイル/ディレクトリの存在確認。 |
| ✅ | **高度な操作** | **ワイルドカード** | `Kill` および `Dir` における `*`, `?` のサポート。 |
| ✅ | | **カレントディレクトリ** | `ChDir` / `CurDir` による仮想的な作業ディレクトリの保持。 | `chdir-curdir.test.ts` |
| N/A | | **永続化** | `localStorage` や `IndexedDB` への保存・復元。（実装予定なし：本エンジンの用途範囲外） |
| ✅ | | **排他制御** | `Lock` / `Unlock` ステートメントのエミュレーション。`Access` / `Lock` 句、範囲指定、省略開始位置、解除、競合時の Error 70 に対応。 | `lock-unlock.test.ts`, `lock-omitted-start.test.ts`, `fs-lock-width.test.ts` |
| ✅ | **互換性** | **バイナリ/テキスト** | `Binary` / `Random` / `Input` / `Output` の主要な入出力に対応。スカラー、型付き配列、固定長 String 配列、UDT 配列、Random レコード長、終端超過 Error 62 を実装済み。多次元配列の左端次元連続順とCP932固定長文字列の物理レイアウトを実Excelで照合済み（`EVAL_LOG.md` XL-001〜XL-007）。 |

### VFS 開発ロードマップ (TODO)

#### 1. コア機能の強化
- ✅ `MemoryFileSystem` にワイルドカードマッチングロジックの実装（`Kill` への統合）。 | `wildcard-kill-dir.test.ts`
- ✅ `MemoryFileSystem` 内部での `cwd` (Current Working Directory) の保持と `ChDir` 対応。 | `chdir-curdir.test.ts`
- ✅ `statSync` / `GetAttr` / `SetAttr` のVBA属性（Read-only, Hidden, Directory 等）を `MemoryFileSystem` で保持・更新。NodeFileSystem は ReadOnly とドットファイルの Hidden を実ファイルから推定し、SetAttr のReadOnlyを chmod に反映する。 | `filesystem.test.ts`

#### 2. ブラウザ環境最適化
- 🔮 **P3（低優先度）** `FileSystem` インターフェースの非同期版（`read`, `write` 等）の検討。現在のEvaluatorは同期APIを前提にしており、ブラウザのメインスレッドをI/Oでブロックする。実装する場合は、FileSystemだけでなくEvaluator・VBA組み込みI/O・呼び出しAPIを非同期化し、既存の同期APIとの互換方針を定める必要がある。
- N/A `IndexedDB` をバックエンドとした `PersistentFileSystem` の実装。（実装予定なし：本エンジンの用途範囲外）

#### 3. テスト環境の改善
- ✅ `filesystem-extra.test.ts` 等のテストコードを、Node.js `fs` 直接参照から `this.fs` (抽象インターフェース) 参照へリファクタリング。（`MemoryFileSystem` + `Evaluator({ fs: vfs })` で対応済み）
- ✅ `VBARunner` の `fs` オプションで複数インスタンス間に単一の `MemoryFileSystem` を注入・共有できるようにした。 | `vba-runner-shared-fs.test.ts`

#### 4. バイナリ操作の極致
- ✅ **`Open For Random Len = N` のレコード位置**: `Len` をハンドルに保持し、`Put` / `Get` のレコード番号を `(record - 1) * Len` バイト位置へ換算。 | `random-record-length.test.ts`
- ✅ **`Put` / `Get` の `Single` / `Double`**: IEEE 754 リトルエンディアン 4 / 8 バイトで直列化・復元。 | `binary-floating-point.test.ts`
- ✅ **`Put` / `Get` の `Date`**: VBA Date シリアル値を IEEE 754 リトルエンディアン 8 バイトで直列化・復元。 | `binary-date.test.ts`
- ✅ **`Put` / `Get` の `Currency`**: 小数点以下4桁でスケールした符号付き64ビット整数をリトルエンディアンで直列化・復元。 | `binary-currency.test.ts`
- ✅ **`Put` / `Get` の固定長 UDT・配列**: 宣言順にスカラー、固定長文字列、数値/Boolean/固定長 String/UDT 配列を連続バイト列として直列化・復元。 | `binary-file-io.test.ts`, `binary-udt-array.test.ts`, `binary-byte-array.test.ts`
- ✅ **可変長 `String` を含む UDT**: UDT要素の2バイト little-endian CP932バイト長ディスクリプターを直列化・復元する。Binary/Randomとも実Excelで `Aあ` → `03 00 41 82 A0`、LOF=9を照合済み（`EVAL_LOG.md` XL-008）。

---

## パーサーの BNF 準拠テスト

MS-VBAL §3.3.5.2 では、識別子が必要な箇所に予約語・コンテキストキーワードを使える場合があります（`unrestricted-name`、`label-name` など）。現在のパーサーでは「識別子チェックが `TokenType.Identifier` のみ」という実装漏れにより、予約語をそれらの位置に置いたときに誤ったパースエラーが発生するケースがあります。

今回の修正例: `Error:` のようにコンテキストキーワードがラベル名に使われた場合、`parseErrorStatement()` に誤って dispatch されて `':'` が expression として解析されていた（`Error:` → `LabelStatement` として正しく認識するよう修正済み）。

### BNF の識別子受け入れ箇所と対応テストの網羅

| 状態 | BNF 位置 | 説明 | 主な対象キーワード | テスト |
|------|----------|------|--------------------|--------|
| ✅ | `label-name:` | ラベル定義（`Error:` など） | `Error`, `Property`, `Class`, `Collection`, `Step`, `Line`, `Kill` etc. | `goto-statement.test.ts` |
| ✅ | `GoTo label-name` | GoTo ターゲット — contextual keyword をラベル名として指定 | 上記同様 | `contextual-kw-labels-goto.test.ts` |
| ✅ | `GoSub label-name` | GoSub ターゲット | 上記同様 | `contextual-kw-labels-goto.test.ts` |
| ✅ | `On Error GoTo label-name` | エラーハンドラーラベル指定（On Error GoTo Error など） | `Error`, `Property` etc. | `contextual-kw-labels-goto.test.ts` |
| ✅ | `On expr GoTo label-list` | On...GoTo の複数ラベルリスト | 上記同様 | `contextual-kw-labels-goto.test.ts` |
| ✅ | `Resume label-name` | Resume ターゲット | 上記同様 | `contextual-kw-labels-goto.test.ts` |
| ✅ | `Dim x As <type-name>` | 型名に予約語を含むケース（例: `As Class`、`As Error` などの contextual keyword） | `Class`, `Collection` etc. | `contextual-kw-labels-goto.test.ts` |
| ✅ | `Event name` / `RaiseEvent name` | イベント宣言・発火に contextual keyword を使用 | `Step`, `Error`, `Text` etc. | `contextual-kw-labels-goto.test.ts` |
| ✅ | `unrestricted-name` 全般 | 変数名・プロシージャ名・パラメーター名に contextual keyword を使用 | 全 contextual keyword | `contextual-keywords.test.ts`, `contextual-keyword-as-identifier.test.ts` |

**実施方針**: BNF で `unrestricted-name` または `label-name` を受け入れる箇所を `parser.ts` 上で列挙し、`CONTEXTUAL_KW` の各トークンを代入したサンプルコードをパースしてエラーが出ないことを確認するテストを `contextual-keywords.test.ts` に追加する。

### BNF vs 再帰下降パーサー 乖離リスト（2026-06-13 調査 / 2026-07-15 BNF 全体監査）

BNF と parser.ts を体系的に比較して判明した未実装・仕様乖離。`^` の左結合性は VBA 実機確認済みで仕様通り。

| 状態 | 優先度 | 項目 | 内容 |
|------|--------|------|------|
| ✅ | P2 | **Mid 代入文** (§5.4.3.8) | `Mid(str, 1, 3) = "XYZ"` の代入文形式。`parseMidStatement` / `evaluateMidStatement` 実装済み。MidB（バイトオフセット）対応済み。 | `mid-statement.test.ts` `mid-statement-length.test.ts` `mid-variants.test.ts` |
| ✅ | P3 | **Def-Directive** (§5.2.2) | `DefInt A-Z` / `DefStr B` 等の暗黙型宣言。レクサー・パーサー・エバリュエーター全対応。`Dim x`（型なし）の初期値・型変換に反映。 | `def-directive.test.ts` |
| ✅ | P3 | **AddressOf のメンバー形式** (§5.6.16.8) | `AddressOf Module.Proc` 形式に対応。`moduleName?: string` フィールドを追加し、評価時は `"Module.Proc"` 文字列を返す。 | `addressof.test.ts` |
| ✅ | P3 | **Type/Enum ボディ内の Rem コメント** (§5.2.3.3/5.2.3.4) | Rem・インラインコメントともに正常パース。バグなし（TODO 記載時点での誤認）。 |
| ✅ | P3 | **Optional パラメータの順序チェック** (§5.3.1.5) | `Sub F(x, Optional y, z)` → Pass1 コンパイルエラー。`validateParameterOrder` を Sub/Function/Event/Declare の全パラメーターリスト解析後に適用。 | `optional-param-order.test.ts` |
| ✅ | P1 | **`On Error GoTo -1`** (§5.4.4.3) | `-1` が `-` + `1` の 2 トークンになりパースエラーだった。`GoTo` 後を EOS まで全結合するよう変更。| `bnf-parser-bug-fixes.test.ts` |
| ✅ | P1 | **`Erase arr1, arr2` 複数要素** (§5.4.3.4) | 1 要素しか解析できなかった。ループ化 + `EraseStatement.name` → `names: Identifier[]`。| `bnf-parser-bug-fixes.test.ts` |
| ✅ | P1 | **`ReDim arr1(5), arr2(10)` 複数変数** (§5.4.3.3) | 1 変数しか解析できなかった。ループ化 + `ReDimStatement` に `declarations: ReDimDeclarator[]` 追加。| `bnf-parser-bug-fixes.test.ts` |
| ✅ | P2 | **`Const A = 1, B = 2` 複数定数** (§5.4.3.2/§5.2.3.2) | 1 定数しか解析できなかった。ループ化 + `ConstDeclaration` に `declarations: ConstDeclaratorItem[]` 追加。| `bnf-parser-bug-fixes.test.ts` |
| ✅ | P1 | **`Open "f.txt" As #1`（For mode 省略）** (§5.4.5.1) | `For` キーワードを必須消費していた。`match` に変更し、省略時 Random モードを既定とした。| `bnf-parser-bug-fixes.test.ts` |
| ✅ | P1 | **ファイル文の `#` なし unmarked-file-number** (§5.4.5.1.1) | Print/Write/Input/Put/Get/Seek/LineInput/Width の 8 文で `#` を必須化していた。`match` に変更し省略可に。| `bnf-parser-bug-fixes.test.ts` |
| ✅ | P2 | **`TypeOf x Is Library.ClassName` ドット修飾型名** (§5.6.7) | `Is` 後のトークンを単一識別子しか受け取れなかった。`.` を消費して多段結合するよう `parsePrimary` を修正。評価時は `__progId__` との照合を追加。 | `bnf-parser-bug-fixes.test.ts`, `typeof-is.test.ts` |
| ✅ | P2 | **`Global` キーワード** (§5.2.3.1) | `Global x As Integer` / `Global Const` / `Global Sub` が認識されなかった。Lexer に `KeywordGlobal` を追加し、Parser で `Public` と同一スコープへマップ。 | `bnf-parser-bug-fixes.test.ts` |
| ✅ | P2 | **`ReDim obj.Arr(n)` / `ReDim .Prop(n)` メンバーアクセス対象** (§5.4.3.3) | `ReDimDeclarator.name` が `Identifier` 固定で、メンバー形式・With 形式が構文エラーになった。`name: Expression` に拡張し、評価器で `MemberExpression`・`ImplicitWithObjectExpression` の各ケースを実装。 | `bnf-parser-bug-fixes.test.ts` |
| ✅ | P2 | **`Next i, j` 複数変数** (§5.4.2.3/5.4.2.4) | 1 変数しか解析できなかった。内側 For が `pendingNextVars` キューに残変数を積み、外側 For が順に消費する方式で実装。ForEach + For の混合 Next も対応。 | `next-multi-var.test.ts` |

### BNF 未修正残件（低優先度 — 修正は任意）

実用コードでほぼ踏まない、またはレガシー専用の未対応構文。対応する場合は上の BNF 準拠テスト表に移動して ✅ にする。

| 状態 | 場所 | 問題 | VBA 例 |
|---|---|---|---|
| ✅ | §5.4.2.1 | `Call .Method arg`（parens なし・With 式 ImplicitWithObject 形式）: 暗黙 `With` メンバーと括弧なし引数を `CallStatement` として受理するよう修正 | `with-call-statement.test.ts` |
| ✅ | §5.4.2.9 | `If x > 0 Then 100`（行番号ジャンプ）: 単行 If の数値節を `GoToStatement` として解釈し、Then / Else の回帰テストを追加 | `if-line-number-goto.test.ts` |
| ✅ | §5.4.2.10/§5.6.9.5 | `Case Is ><5` / `><` `=<` `=>` 代替比較演算子: 字句解析で標準比較演算子へ正規化し、式と `Case Is` を回帰テスト化 | `alternate-comparison-operators.test.ts` |
| ✅ | §5.4.2.12 | `Go To`（2語形式）: `Go` と `To` を連続キーワードとして受理し、名前・数値ラベルへのジャンプを回帰テスト化 | `go-to-statement.test.ts` |
| ✅ | §5.4.2.14 | `Go Sub`（2語形式）: `Go` と `Sub` を連続キーワードとして受理し、名前・数値ラベルを回帰テスト化 | `go-sub-statement.test.ts` |
| ✅ | §5.4.3.1 | `Dim Shared x As Integer`（VBA6 方言）: `Shared` 修飾子を保持して通常の `Dim` として受理 | `dim-shared.test.ts` |
| ✅ | §5.4.5.1 | `Open "f.txt" For Random Access Read Shared As #1`（`Shared` スタンドアロン）: Access 指定と併用して受理 | `open-shared-lock.test.ts` |
| ✅ | §5.4.5.4/5 | `Lock #1, To 100`（start 省略の `To end` 形式）: start を 1 として扱い、対になる `Unlock` も受理 | `lock-omitted-start.test.ts` |
| ✅ | §5.4.5.6 | `Line Input #1, arr(i)`（variable が配列要素）: 代入先を識別子に限定せず配列要素を受理 | `line-input-array-element.test.ts` |
| ✅ | §5.2.3.3 | UDT メンバー名として `End`: `End Type` の組だけを終端とし、メンバー名を受理 | `udt-end-member.test.ts` |
| ✅ | §5.2.4.1.2 | `Implements SomeLib.IInterface`（ドット修飾名）: ドット区切りのインターフェース名を受理 | `implements-qualified-name.test.ts` |
| ✅ | §5.3.1.3 | `Function Foo$()` 等 TYPED-NAME サフィックス付き宣言名: 宣言時にサフィックスを戻り型へ反映し、通常名 `Foo()` で呼び出せる | `typed-procedure-name.test.ts` |
| ✅ | §5.3.1.5 | `ByVal Optional x As Integer`（ByVal が先）: パラメーター修飾子を順不同で受理 | `parameter-modifier-order.test.ts` |
| ✅ | §5.6.5 | `1.5D10`（VBA Double `D` 指数表記）: **実 VBA 実機確認済み（2026-07-19）**の `x = 1.5D+10` / `y = 2D5`（ともに `Double`）に合わせ、レキサーが `D`/`d` を指数部として受理し、パーサーで JavaScript の `e` へ正規化して数値化するよう修正。`1.5D+10` / `2D5` / `1d-2` と `TypeName` の回帰テストを追加 | 科学技術計算 |
| ✅ | §5.6.9.5 | 連鎖比較演算子: 比較演算子を同一優先度・左結合として解析し、`a = b < c` を `(a = b) < c` と評価 | `comparison-operator-precedence.test.ts` |
| ✅ | §5.6.13.1 | `Foo(ByVal x)`（呼び出し側 ByVal）: 呼び出し引数を値渡しとして受理し、ByRef 書き戻しを抑止 | `call-byval-argument.test.ts` |
| ✅ | §5.6.15 | `With dict: !Key = v`（行頭 `!`）: With の辞書アクセスを代入・参照とも受理 | `with-dictionary-bang.test.ts` |
| ✅ | §5.6.16.7 | `Dim x As A.B.C`（type-expression の2段階以上ドット修飾）: 多段の型修飾名を保持して受理 | `multi-qualified-type-name.test.ts` |
| ✅ | §5.6.16.8 | `AddressOf NS.Module.Proc`（多段階ドット）: 最終要素を手続き名、前段をモジュール名として保持 | `address-of-qualified-name.test.ts` |

---

## VBA ランタイム挙動

過去の仕様バグの詳細な症状・原因・修正履歴は [VBA_ENGINE_BUG_HISTORY.md](../bug-records/VBA_ENGINE_BUG_HISTORY.md) に記録する。

仕様書本文には書かれているが、個別セクションを持たないため見落とされやすい **暗黙のランタイム挙動** の実装状況をトラッキングするセクションです。チェックなしの項目は **未実装 または 未検証** を意味します。実装する際は「ロードマップ」の手順に従い、VBA ベテランが書きそうな代表コードをテストに含めてください。

### オブジェクトのライフサイクルと参照

- ✅ **Auto-Instantiation (`Dim x As New ClassName`)**: 遅延インスタンス化 + `Set = Nothing` 後の自動再生成 + `Is Nothing` 常時 False
  - ✅ **Auto-Instanceを関数/メソッド引数として渡す**: 未アクセスのAuto-Instanceを引数に渡す際にプレースホルダーのまま渡されるバグを修正 | テスト: `auto-instance-args.test.ts`
    - 修正箇所: JS組み込み関数引数 (L4611)、VBAクラスメソッド引数 (L4698/L4704)、JSオブジェクトメソッド引数 (L4727)、汎用フォールバック (L4752)、デフォルトプロパティ引数 (L4641)
    - 実害バグ: `CallByName body, "Method", VbMethod, autoInstance`（`__vbaClass__` 直接参照で失敗）および `d.Add b, val` でbをキーに使う場合（後でbが解決されると`Exists`がFalse）
- ✅ **`Dim x As ClassName`（New なし）のデフォルト値**: `Nothing` 初期化 | `on-error-nothing.test.ts`
- ✅ **Default Property / Default Member**: `Range("A1") = 10` のような暗黙の `.Value` 解決、Collection の `Item` 暗黙呼び出し等
  - ✅ 包括的なテストスイート作成 (`default-property.test.ts`, 15テスト全パス)
  - ✅ 基盤研究: VBA 仕様書からの解析、実装パターン検討
  - ✅ 実装完了: (1) obj = value -> obj.Value = value (Test 14パス)
  - ✅ 実装完了: (2) obj(args) -> obj.Item(args) (Test 12-13パス)
  - ✅ 実装完了: (3) result = obj -> obj.Value (Test 15パス) - 値コンテキストでの暗黙 Value getter
  - ✅ 修正: Test 7 ByRef パラメーター - implicit Value getter の実装で副次的に修正
  - ✅ **非 `__vbaClass__` モックオブジェクトのデフォルトプロパティ**: `x = ws.Range("A1")` で MockRange の `.Value` が自動抽出される
    - opt-in 方式: モックオブジェクトに `__vbaDefault__ = true` と `Value` getter/setter を実装することで有効化
    - 読み書き両方 `Value` プロパティ経由（対称）。evaluator は `resolveObjectMemberKey(obj, 'value')` で解決
    - `VbaDate` / `VbaBoolean` / `VbaErrorValue` 等の内部型は `__vbaDefault__` を持たないため誤抽出しない
    - `MockRange` はすでに対応済み (`__vbaDefault__ = true`, `Value` getter/setter 実装)
    - テスト: `default-property-noncls.test.ts`
- ✅ **`WithEvents` の再代入・`Set ... = Nothing` による購読解除**: フィールドごとにイベントコールバックを記録し、再代入と Nothing 時に古いイベントソースから解除する。別変数がソースを保持していても、解除後はハンドラーへ通知されない。 | `raiseevent.test.ts`
- ⚠️ **恒久的制限: `DoEvents` + `Sleep` による外部イベント待ち**: `DoEvents` 自体は呼び出せるが直ちに戻るだけで、Evaluatorの実行を中断してイベントループへ制御を返さない。`Declare Sub Sleep ...` などの外部APIはスタブとして即時に戻るため、次の動作はできない。
  - 別スレッド・タイマー・UI操作から実行中のVBAへイベントを届ける
  - `DoEvents` でイベントを処理し、状態が変わるまで `Sleep` して待つ
  - `wait/notify`、キャンセル、タイムアウトを伴うイベント待ちを再現する
  これを解決すると、Evaluatorが実行を一時停止して外部イベントを受信し、イベント処理後に同じVBAスタックを再開する非同期実行モデルが必要になる。現在の同期Evaluatorの設計範囲外であり、恒久的制限として扱う。
- ✅ **循環参照時の `Set = Nothing` 挙動**: 強制クリアと Class_Terminate の呼び出し順 | テスト: `circular-reference-terminate.test.ts`, `Circular/TerminateTest.bas` (VBA: `Circular/Helper.cls`, `Circular/RefA.cls`, `Circular/RefB.cls`, `Circular/TerminateTest.bas`)
- ✅ **`Me` キーワードの完全対応**: クラスモジュール内での全コンテキスト | `me-keyword.test.ts`
- ✅ **`Implements` インターフェース呼び出し**: `obj.Speak` → `IAnimal_Speak` のインターフェースディスパッチ | テスト: `implements-dispatch.test.ts`

### 型変換と暗黙のキャスト

- ✅ **`Variant` 型の値伝播**: `Variant + Number` の結果型の自動拡張 (`variant-propagation.test.ts`)
- ✅ **`Empty` / `Null` / `Nothing` / `Missing` の演算子挙動**: `Null + 5 = Null`、`Empty + 5 = 5` などの伝播規則 (`null-empty-propagation.test.ts`)
- ✅ **`String` ↔ 数値の暗黙変換**: `"5" + 3 = 8`、`"5" & 3 = "53"` の区別 (`string-numeric-coercion.test.ts`)
- ✅ **`Boolean` の代入時の型強制**: 非 0 数値・文字列 "True"/"False"・数値文字列・Empty を Boolean に変換、Null は Error 94 (`boolean-coercion.test.ts`)
- ✅ **`If` 条件式の truthy 判定**: 数値・Boolean・文字列（"True"/"False"/数値文字列→変換、それ以外 Type mismatch）・Null(Error 94)・Empty(False)。Boolean 算術（True=-1）・Not ビット反転・比較演算子結果が vbaTrue(-1)/vbaFalse(0)。文字列変換は §5.6.9（条件式は §6.1.2.3 の value coercion を適用）＋ §6.1.2.3.1.1（CBool 変換規則）が根拠 (`if-condition-truthy.test.ts`)
- ✅ **Integer/Long のオーバーフロー検出**: Error 6 を正しく発生させる (Byte/Integer/Long/Currency/Single)
- ✅ **`LongPtr` 型のサポート**: 64bit 環境で `LongLong` と同等（`bigint`、VarType=20）。`Dim x As LongPtr`・`CLngPtr()` 対応 | `clnglng.test.ts`
- ✅ **関数戻り値への Let-coercion**: `Function F() As Long` に対して `F = 3.7` → 4 に丸め、`F = "hello"` → Type Mismatch (Error 13) など宣言型に応じた強制変換を実装。原因は `callProcedure`/`evaluateCallExpression` の両パスで `setVariableType` が呼ばれていなかったため。 | `function-return-coercion.test.ts`
- ✅ **`Decimal` 型の精度**: 28 桁の固定精度サポート。`VbaDecimal` 内部を `mantissa: bigint` + `scale: number`（96-bit）に刷新。`CDec(string)` は float 非経由で直接パース。`CDec(1)/CDec(3)` = 28桁の "0.333...3" | `decimal-precision.test.ts`, `decimal-bigint.test.ts`
- ✅ **`Currency` 型の精度**: 64-bit 整数 / 10000 の固定小数 | `currency-precision.test.ts`

### Variant 配列とコレクション

- ✅ **2 次元 Variant 配列の `For Each` 挙動**: 行優先 vs 列優先の反復順 | `2d-array-foreach.test.ts`
- ✅ **Array() で作る Variant 配列の境界**: `LBound = 0`、明示的 `Option Base` の効果 | `array-lbound-defaults.test.ts`
- ✅ **`ReDim Preserve` の多次元配列での制約**: 最終次元のみ変更可 | `redim-preserve-multidim.test.ts`
- ✅ **UDT 配列の `ReDim` 時の要素初期化**: `Dim d() As MyType` + `ReDim d(0 To N)` で各要素が UDT インスタンスではなく `0` になるバグを修正 | `udt-type.test.ts`
- ✅ **`ReDim` 後の型付き配列要素の型強制**: `Dim d() As Integer` の `ReDim` / `ReDim Preserve` 後も、要素代入を Integer として強制変換する。 | `array-functions.test.ts`
- ✅ **クラスフィールド配列の型強制**: `Private d() As Integer` の `ReDim` 後も、通常のローカル配列と同じく要素代入を宣言型へ強制変換する。 | `class-module.test.ts`
- ✅ **`Byte` 配列への文字列の暗黙代入**（`Dim b() As Byte : b = "AB"`）
  - 実 VBA 実機確認（2026-07-19）: 代入は合法で、`UBound(b) = 3`（1文字 = 2バイト）。中身は UTF-16LE のリトルエンディアン生バイト列（`"AB"` → `b(0)=&H41('A'下位), b(1)=&H00, b(2)=&H42('B'下位), b(3)=&H00`）
  - 評価 #37 で修正: 文字列代入時に対象が `Byte()` なら UTF-16LE のコード単位へ展開し、下限 0 の動的配列として保持する。`"Aあ"` → `41 00 42 30` を回帰テスト化。

### エラーハンドリング

- ✅ **`On Error Resume Next` 下の暗黙の `Err.Clear`**: 正常文実行時のクリア（または非クリア）タイミング | `err-clear-timing.test.ts`
- ✅ **`Err.Source` の既定値と更新**: `Err.Raise` の Source 省略時は既定値を設定し、暗黙ランタイムエラーでは直前のカスタム Source を引き継がない。 | `error_handling.test.ts`（評価 #181）
- ✅ **`Resume` の対象決定**: エラー発生点・Resume Next・Resume <label> の正確な制御フロー | `resume-statement-target.test.ts`
- ✅ **エラーハンドラー内での再帰的なエラー発生**: スタックフレームのリセット規則 | `recursive-error-handling.test.ts`
- ✅ **`Erl` 関数**: エラー発生時点で最後に通過した数値行ラベルを返す。数値行ラベルがない VBA ソースでは仕様どおり 0 を返す。 | `erl-function.test.ts`

### プロシージャ呼び出しの細部

- ✅ **モジュール修飾付きプロシージャ呼び出し**: `ModuleName.ProcedureName()` による複数ファイル間での同名関数の区別
  - 実装: `Environment.setProcedureWithModule()` / `getProcedureFromModule()` でモジュール名付き登録・検索
  - VBA 名前解決優先度（修飾なし呼び出し）:
    1. グローバルスコープ（モジュール名なし）← 最優先
    2. 単一モジュール内のプロシージャ（曖昧でない）
    3. 複数モジュール → 曖昧性エラー（修飾必須）
  - 曖昧性検出: 複数モジュールに同名プロシージャがある場合、実行時に詳細エラーを発生
  - テスト: `module-qualified-calls.test.ts`, `ambiguous-procedure-call.test.ts`, `evaluator-scoping.test.ts`
- ✅ **モジュール修飾付き変数/定数アクセス**: `Module1.A` 形式でモジュールレベルの変数・定数を参照
  - 実装: Const は module-qualified キー (`module1:a`) で格納（不変なので複製コピーで同名競合も区別可）。変数は `moduleVarRegistry` に登録し参照時は非修飾名で引く
  - `evaluateMemberExpression` でオブジェクト評価前に台帳チェック（`Environment.get` の暗黙ゼロ初期化による誤検知を防止）
  - 制限事項: 同名のモジュールレベル変数（Const でない）が複数モジュールに存在する場合、最後の代入が勝つ（区別不可）
  - テスト: `module-qualified-access.test.ts`
- ✅ **ByRef での文字列・配列・オブジェクトの参照保持**: 文字列・数値・Boolean・配列・オブジェクト全て正常動作 | `byref-reference-preservation.test.ts`
- ✅ **ParamArray の境界ケース**: 0 個渡し、配列を 1 つだけ渡したときの展開規則、ByRef semantics (spec §5.3.1.5) | `paramarray-edge-cases.test.ts`
- ✅ **Optional パラメーターの IsMissing 判定**: デフォルト値ありと未指定の区別
- ✅ **Property Get/Let/Set の解決順序**: 同名で混在した場合の優先度 | `property-resolution-order.test.ts`
- ✅ **暗黙の Let（`Call` なしの呼び出し）**: 戻り値が破棄される / されない場面 | `implicit-let.test.ts`
- ✅ **引数の数の検証**: プロシージャ呼び出し時に引数の個数が定義と異なる場合にエラーを発生させる (Error 450 / 449) | テスト: `ArgCountTest.bas`
  - 制限事項: VBA 本来はコンパイルエラーだが、本インタープリターでは実行時エラーとして発生させる
- ✅ **位置引数の省略（`Func(a,,c)`）**: `,,` で中間引数を省略した場合に Optional のデフォルト値が適用される | テスト: `missing-arg.test.ts`
  - 修正: 省略スロット（`MissingArgument` ノード）が `positionalArgs` に含まれていてもデフォルト値分岐に落ちるよう修正
  - 省略された必須パラメーターは Error 449 として拒否する。修正の経緯は [VBA_ENGINE_BUG_HISTORY.md](../bug-records/VBA_ENGINE_BUG_HISTORY.md) を参照。 | テスト: `missing-arg.test.ts`

### 数値・日付の演算

- ✅ **`\`（整数除算）の VBA 厳密仕様**: 引数の自動型変換（Variant → Integer/Long）と余り処理
- ✅ **`Mod` 演算子のオーバーロード**: 浮動小数 vs 整数の挙動差
- ✅ **`#1/1/2024#` 型 Date リテラルの月日順序**: ロケールに依存しない解釈の保証 | `date-literal-parsing.test.ts`
- ✅ **`DateAdd("m", 1, ...)` 等の月末ロールオーバー規則**: Jan31 + 1 month = Feb28/29 | `dateadd-monthend-rollover.test.ts`

### 文字列処理

- ✅ **`Like` 演算子の文字クラス `[a-z]` 詳細**: 範囲外文字・特殊エスケープの扱い全て対応済み | `like-charclass.test.ts`
- ✅ **`Mid$` 代入文の長さ規則**: 元文字列長を超える代入の挙動 | `mid-statement-length.test.ts`
- ✅ **`StrComp` の Option Compare 影響**: モジュールレベル設定の伝播 | `strcomp-option-compare.test.ts`
- ✅ **ANSI/Unicode 切り替え**: `StrConv(..., vbFromUnicode)` 等の挙動 | `strconv-unicode.test.ts`

### Late Binding 周辺

- ✅ **`CreateObject(progId)` 経由のオブジェクト呼び出し**: 動的ディスパッチ
- ✅ **存在しないメソッド呼び出しのエラー詳細**: VBA エラー番号 438 を厳密に再現 — クラスインスタンス・CreateObject 返り値いずれも 438 | `ObjectErrorTest.bas`
- ✅ **`IDispatch` 風のプロパティ get/set 自動振り分け**: 引数の有無でメソッド/プロパティ自動判定 | `idispatch-property-dispatch.test.ts`

---

## エンジンの改善

過去の仕様バグの詳細な症状・原因・修正履歴は [VBA_ENGINE_BUG_HISTORY.md](../bug-records/VBA_ENGINE_BUG_HISTORY.md) に記録する。

### 組み込み関数の引数メタデータ化（auto-call・引数数検証・名前付き引数）

- ✅ **組み込み関数に VBA パラメーターメタデータ（`BuiltinParamSpec[]`）を持たせる仕組みを導入**
  - 経緯: 当初は「必須引数0個なら括弧無しで自動呼び出し」判定（`__vbaAutoCall__`）の手動フラグ管理が個別登録に依存しており、`Rnd`/`DoEvents`/`FreeFile`/`GetObject` でフラグ漏れバグが発生していた（2026-06-21 修正）。これを機に、ユーザー定義 Sub/Function 同様の「引数の数の検証」「名前付き引数（`:=`）の解決」も組み込み関数に持たせる方向に拡張した。
  - 実装: `registerBuiltin(name, fn, params: BuiltinParamSpec[], variants?)` — `__vbaParamSpec__` を関数に付与する。呼び出し時は `resolveCallArgs`（`evaluator.ts`、グローバル組み込み関数呼び出し・`VBA.Func(...)` 修飾呼び出し・`obj.Method(...)` メソッド呼び出し・式の結果を呼ぶ汎用フォールバックの4箇所で共有）が `__vbaParamSpec__` の有無を見て、検証・名前解決を行うか今までの単純な位置引数評価のみで済ませるかを切り替える（未移行の関数は無変化）。
  - 追記（2026-06-21）: `__vbaAutoCall__` フラグ自体も廃止した。`registerBuiltin`/`registerOverloadedBuiltin` で必ず `__vbaParamSpec__`/`__vbaOverloads__` が付与されるようになったため、「必須引数0個か」を事前計算してキャッシュする別フラグは冗長だった。括弧無し参照時（`evaluator.ts` の `Identifier` 評価）は `isAutoCallable(fn)` がその場で `__vbaParamSpec__`/`__vbaOverloads__` から算出する形に置き換えた。`spy()` のメタデータ伝播からも `__vbaAutoCall__` のコピーを削除。
  - `InStr`/`InStrB` のように引数の個数で意味が変わる不規則な組み込み関数（`Start` が先頭にある Optional 引数）は `registerOverloadedBuiltin(name, fn, overloads: BuiltinOverload[])` で対応。VBA 自体にはないオーバーロード機構をエンジン内部だけで提供する。位置引数のみの呼び出しは個数検証のみ行い関数本体（既存の `args.length`/`typeof args[0]` 判定）はそのまま使う。名前付き引数を使った場合のみ該当オーバーロードを選んで引数を再構築する。
  - 移行済み: `Now`/`Date`/`Time`/`Timer`/`Rnd`/`Randomize`/`FreeFile`/`DoEvents`/`Error`/`GetObject`（Batch 1）、`InStr`/`InStrB`（Batch 3 先行）。
    Batch 2（単純な1必須引数のみ、または必須1+末尾Optional1）51個: 型変換 `CByte`/`CInt`/`CLng`/`CSng`/`CDbl`/`CDate`/`CVDate`/`CDec`/`CCur`/`CLngLng`/`CLngPtr`/`CStr`/`CBool`/`CVar`/`CVErr`/`Hex`/`Oct`/`Val`、情報関数 `IsEmpty`/`IsMissing`/`IsNumeric`/`IsDate`/`IsObject`/`IsError`/`IsNull`/`IsArray`/`VarType`/`TypeName`、数学関数 `Abs`/`Atn`/`Cos`/`Exp`/`Int`/`Fix`/`Log`/`Round`/`Sgn`/`Sin`/`Sqr`/`Tan`、文字列関数 `Asc`/`AscW`/`Chr`/`ChrW`/`LCase`/`Str`/`UCase`/`Len`/`LTrim`/`RTrim`/`Trim`/`Space`/`StrReverse`。
    Batch 3（複数引数・Optional引数を持つ関数）: 文字列 `InStrRev`/`Left`/`Right`/`Mid`/`String`/`Split`/`Join`/`Replace`/`StrComp`/`StrConv`/`Filter`/`LeftB`/`RightB`/`MidB`/`Format`、対話関数 `Shell`/`MsgBox`/`InputBox`/`AppActivate`/`SendKeys`、財務関数 `FV`/`PV`/`Pmt`/`NPer`/`Rate`/`SLN`/`SYD`/`DDB`/`IRR`/`MIRR`/`NPV`/`IPmt`/`PPmt`。
    Batch 4（日時 stdlib・ファイル操作・レジストリ・残りの雑多な関数）: 日時 `Year`/`Month`/`Day`/`Hour`/`Minute`/`Second`/`DateSerial`/`TimeSerial`/`Weekday`/`DateAdd`/`DateDiff`/`DatePart`/`DateValue`/`TimeValue`/`MonthName`/`WeekdayName`、ファイル操作 `EOF`/`LOF`/`Loc`/`Seek`/`FileAttr`/`ChDrive`/`SetAttr`/`FileDateTime`/`CurDir`/`Dir`/`FileCopy`/`Kill`/`MkDir`/`RmDir`/`ChDir`/`FileLen`、レジストリ `SaveSetting`/`GetSetting`/`GetAllSettings`/`DeleteSetting`、雑多 `Environ`/`VarPtr`/`StrPtr`/`ObjPtr`/`CreateObject`/`IIf`/`Choose`/`Switch`/`Array`/`LBound`/`UBound`/`CallByName`。`Choose`/`Switch`/`Array`/`CallByName` の可変長引数部分は `BuiltinParamSpec.isParamArray` で表現し、引数数検証をスキップして既存の可変長挙動をそのまま維持。
    これで組み込み関数登録箇所の段階移行は完了（モックオブジェクトメソッドは別件、下記）。
    テスト: `builtin-arg-metadata.test.ts`。
  - モックオブジェクトメソッドへの適用（`MockWorksheet.Range`/`MockApplication.Range` の2引数対応）: `Range(Cell1, Cell2)` — 2つの角セル（Address 文字列または `Range`/`Cells()` の戻り値）を結ぶ矩形を返す新機能を実装。`registerOverloadedBuiltin` と同じ仕組みを使い、`MockWorksheet.prototype.Range`/`MockApplication.prototype.Range` に直接 `__vbaOverloads__`（`[{Cell1}]`/`[{Cell1,Cell2}]`）を付与。`resolveCallArgs` が `obj.Method(...)` 呼び出し・Tier 6 (`defaultBindingObject`) 呼び出しの両方でこれを解釈するため、`Range("A1","C3")` の位置引数はもちろん、名前付き引数 `Range(Cell2:="B2", Cell1:="A1")`（順序非依存）や引数過多エラー（450）も組み込み関数と同様に動作する。
    テスト: `mock-worksheet-address.test.ts`。

### テストランナー (`VBARunner`) の改善

この節の本文中にある「修正前の現状」欄の ❌ は、修正前の挙動を
記録した履歴であり、現在の未実装項目を示すものではない。現在の残件は
各仕様セクションの状態表（および ⚠️ / 🔮 の項目）を参照する。

- ✅ **`VBARunner()` 引数なし・`null` 対応**: コンストラクターの `pathOrDir` を `string | null = null` に変更し、空文字・`null`・引数省略すべてで「ファイルなし」の空環境を作成できるよう修正
  - 修正前は `fs.statSync('')` が ENOENT を投げるため、Unix の `/dev/null` 回避策が必要だった
  - テスト: `time-mocking.test.ts`（test 8 で `new VBARunner()` を使用）

### 条件付きコンパイル (§3.4)

- ✅ **`#If` / `#ElseIf` / `#Else` / `#End If` / `#Const` 対応**: `src/engine/preprocessor.ts` でソーステキストを Lexer 前処理 | `conditional-compilation.test.ts`
  - デフォルト定数: `VBA7=0`, `Win64=0`, `Win32=-1`, `Mac=0`（ホストアプリ定義の project-level constants）
  - `VBARunner` コンストラクターの `config.compilerConstants` で上書き可能
  - 未定義シンボルは `0`（falsy）として扱う（VBA 仕様に準拠）
  - `#Const` はすべてのブロック（excluded block も含む）で処理される（仕様 §3.4.1）
  - ネスト・`#EndIf`（スペースなし）・`And`/`Or`/`Not` 演算子すべて対応

### Parser / Evaluator バグ修正履歴

過去の仕様バグの症状・原因・修正内容は
[VBAエンジンバグ履歴](../bug-records/VBA_ENGINE_BUG_HISTORY.md)を参照する。
この文書では、現在の仕様準拠状態と残る制限だけを管理する。

### Parser の拡張機能

- ✅ **Parser に `parseAsClass` パラメーターを追加**: `.cls` ファイルやプログラム的にクラスとしてパースすべきコードの指定 | `parse-as-class.test.ts`
  - `new Parser(tokens, { parseAsClass: 'ClassName' })` でクラスボディとしてパース
  - `parseClassDeclaration` のボディ解析を `parseClassBody(name, untilEndClass)` に抽出し再利用
  - `test-libs/test-runner.ts` の文字列ラップハックを削除し `parseAsClass` オプションに統一

- ✅ **Parser の EOS 検証を体系化する**（構文精度の根本課題）
  - 背景: MS-VBAL §5.4.1 `statement-block = *(block-statement EOS)` / `EOS = *(EOL / ":")` によれば、各文の後には必ず EOS（改行またはコロン）が必要。
  - 実装: `parseStatement(checkEOS = true)` にデフォルト引数を追加。stmt が非 null かつ `!isAtTerminator()` のとき ParseError を投げる。インライン If の2ループのみ `parseStatement(false)` を渡す（`Else` トークンが EOS の役割を担うため）。
  - 効果: `Return Foo`・`Exit Sub Bar`・`Exit For Extra` 等の余分トークンが全ブロックループで自動的に検出される。`End Foo`・`Stop Foo` の個別チェック（既存）は冗長になるが無害。
  - テスト: `eos-verification.test.ts` | 2026-06-06

- ✅ **Lexer のキーワード分類を仕様書 §3.3.5.2 のカテゴリに整理する**（可読性・保守性）
  - 現状: 全キーワードが `Keyword*` トークンのフラットな羅列で、仕様上の種別（`statement-keyword` / `marker-keyword` / `operator-identifier` / contextual）が区別されていない
  - contextual キーワードの誤予約語化（`Append`・`Output` 等が変数名に使えない問題）は Parser の `CONTEXTUAL_KW` Set 化により軽減済み。新規追加も Set に1行追加するだけ
  - 2025-05-30: `Class / Collection / Error / Property` の未登録を修正、`isIdentifier()` ヘルパー導入、`parseVarDeclaration` / `parseConstDeclaration` / `parseProcedureDeclaration` のチェックを統一。詳細は `docs/internals/IDENTIFIER_CLASSIFICATION.md` を参照
  - 2026-06-06: `parseEnumDeclaration` / `parseTypeDeclaration` / `parseForStatement` / `parseForEachStatementBody` の変数名・型名チェックを `isIdentifier()` に統一。`Next <var>` のチェックも同様に修正 | `contextual-kw-for-enum-type.test.ts`
  - 将来検討: キーワード補完実装時に contextual keyword を候補から除外する仕組みを補完側へ追加する
  - `VBA_KEYWORD_CATEGORIES` に仕様カテゴリをメタデータとして追加。`TokenType` と既存のParser動作は変更せず、`COMPAT_KW_EXPR` とも分離した

- ✅ **予約語 + 型接尾辞（`dim$`・`for$` 等）が変数名として通ってしまう**（バグ）
  - 修正: Lexer でキーワード照合前に型接尾辞を除去（`lowerBase`）することで `dim$` → `KeywordDim` と正しく識別
  - ただし contextual keyword + `$`（`append$` 等）は IDENTIFIER なので合法 | `lexer-column.test.ts`

### エンジン内部の構造改善

- ✅ **型変換（Coercion）ロジックの一元化**: `coerce.ts` を新設し `vbaToNumber()`/`vbaToString()`/`vbaToBoolean()`/`vbaToDisplayString()`/`vbaRound()` を集約する
  - VBA 型定義は `vba-types.ts` に抽出。`evaluator.ts` の各変換メソッドは委譲形に置き換え、re-export で後方互換を維持 | `coerce.ts`, `vba-types.ts`

- ✅ **`registerStandardLibrary` の分割 + `$` suffix 宣言的登録**: 940 行の単一メソッドを 11 の private メソッドに分割し、`envSet(name, fn, ['$'])` ヘルパーで 21 件の `$` variant 二重登録を一元化 | `evaluator.ts`
  - 現状: `Left`/`Left$` を別々にif分岐で手動登録しており、新しい組み込み関数追加のたびに `$` 版を忘れるリスクがある

- ✅ **エラー番号の一元管理**: `vba-errors.ts` を新設し、全エラー番号・メッセージ・カテゴリを辞書として集約する
  - `VbaErrorCode` 定数・`VBA_ERROR_MESSAGES` 辞書・`throwVbaError` ヘルパーを実装。evaluator.ts / coerce.ts / vba-types.ts のマジックナンバーをすべて定数に置き換え | `vba-errors.ts`

- ✅ **`Option Explicit` 検証の実装**: `Option Explicit` 宣言がある場合、未宣言変数を静的解析 + 実行時エラーで検出
  - 実装: `src/engine/option-explicit-checker.ts`; AST を2パスで解析し `program.diagnostics` に追記、違反SubはLSP破線＋呼び出し時にVBAエラー1で停止
  - テスト: `option-explicit.test.ts`
  - ✅ **マルチモジュール 2nd pass チェック強化**: `resolveIdentifiers()` 呼び出し時に全モジュール名を `checkOptionExplicit` へ渡し、コール式の bare identifier オブジェクト（`undeclaredObj.Method()` 形式）を精密に検出。既知モジュール名のみスキップ、それ以外は未宣言エラー | `option-explicit.test.ts`（Test 11-12）

- ✅ **`Identifier` AST ノードへの `loc` 付与 + スコープ対応シンボルテーブル**: 宣言位置の Identifier に `loc` を付与し、LSP のシンボル参照をスコープ対応に刷新
  - `parser.ts`: `makeIdentifier(token)` ヘルパー追加；Dim・Sub 名・For 変数・Const など全宣言位置に正確な `loc`
  - `symbol-table.ts`: フラット Map → `ScopedSymbolTable`（手続きスコープ別 `localSymbols`）+ `lookupSymbol()`（内側優先解決）
  - 効果: 同名ローカル変数が別 Sub にある場合でも「定義へ移動」「参照検索」「ホバー」「リネーム」が正しいスコープを対象にする
  - `extension.ts` / `code-lens-provider.ts`: `findReferences` コマンドの引数バグ修正（`procName` → `line, character`）

- ✅ **`precheckProc` の AST チェックを 1 パス走査に統合**
  - `collectPrecheckFindings` の共通深さ優先走査で、Sub-as-value、未定義呼び出し、定数配列境界、重複宣言、ラベル、呼び出し引数数を収集する
  - `If` / `For` / `For Each` / `Do` / `While` / `With` / `Select Case` の全分岐・ネストを走査し、従来のエラー優先順位と行番号を維持して走査後に報告する
  - `Option Explicit` の呼び出し時再評価は既存の責務として維持する。旧6チェッカーの独立走査は削除した
  - `vba_compile_error.test.ts`、配列境界、GoTo/GoSub、名前空間、TypeName、および全体テスト（TypeScript 330ファイル / VBA 14ファイル）で回帰確認済み

- ✅ **`precheckProc` のコンパイルエラーメッセージ "Compile error:" プレフィックスをキャッチ側で一元付与**
  - 各チェックメソッドは `throwPrecheckError`（プレフィックスなし・`_precheckRaw` マーカー付き）でスロー
  - `precheckProc` の catch ブロックで "Compile error:" プレフィックスと行番号を一元付与

### VBA 仕様制約の検証

- ✅ **モジュール名の長さ検証（31 文字制限）**: MS-VBAL §5.2 で定義されたモジュール名の最大長を実行時に検証
  - 仕様: モジュール名は最大 31 文字（ファイル名またはVB_Name属性）
  - 実装: `Evaluator.setSourceModule()` でモジュール名長をチェック、超過時にエラー発生
  - テスト: `module-name-length.test.ts`, `module-name-length-integration.test.ts`

- ✅ **§5.6.10 Tier 6 — 外部型ライブラリのメンバーを修飾なしで解決できる**
  - **仕様**: MS-VBAL §5.6.10 では修飾なし識別子の名前解決を 6 段階の Tier で定義している。
    Tier 1（ローカル）→ Tier 2（同モジュール）→ Tier 3（プロジェクト名）→ Tier 4（他モジュール Public）→ Tier 5（参照プロジェクト名）→ **Tier 6（参照プロジェクト内 Public メンバー）**
  - **Tier 6 の具体例**: Excel VBA では参照設定により Excel 型ライブラリが Tier 6 に入る。
    `Range("A1")` が修飾なしで呼べるのは Excel 型ライブラリが `Application` の `Range` メンバーを
    Tier 6 に公開しているため。Default Member 機構により `Range("A1")` → `Application.ActiveSheet.Range.Item("A1")` に解決される
  - **型文脈との関係**: `Dim r As Range` の `Range`（型名前空間）と `Range("A1")` の `Range`（値名前空間）は
    同じ Tier 6 で解決されるが、型名前空間と値名前空間は別扱いのため名前衝突は起きない
  - **実装**: `evalVBAModules` の `defaultBindingObject` オプションに `MockApplication` 等を渡すことで
    `Range("A1")`、`Cells(row, col)`、`ActiveSheet`、`Sheets("name")` 等を修飾なしで解決できる。
    `Option Explicit` 環境でも `defaultBindingObject` のメンバーは暗黙変数扱いにならず適切に解決される。
  - テスト: `tier6-namespace.test.ts` | `tests/spec/tier6-namespace.test.ts`

- ✅ **静的 vs 動的名前解決の区別**: プロシージャ呼び出し時のエラー検出タイミング
  - 実VBA動作の違い：
    - `unknownProc()` → **Pass2 コンパイルエラー**（Sub/Function が定義されていません）
    - `unknownModule.unknownProc()` with Option Explicit → **Pass2 コンパイルエラー**（変数が定義されていません: unknownModule）
    - `unknownModule.unknownProc()` without Option Explicit → **実行時エラー 424**（Dim が暗黙挿入されオブジェクト未設定）
  - 実装:
    - `collectUndefinedProcCalls`（`option-explicit-checker.ts`）: 非修飾 bare Identifier callee を Pass2 で静的検証
    - Option Explicit 違反を `callProcedure` 実行時エラー（遅延）→ `resolveIdentifiers`（Pass2）で即時 throw に変更
    - `defaultBindingObject`（Tier 6）設定済みならメンバーを既知として除外（Excel 互換維持）
    - `evalVBAModules` に `defaultBindingObject` オプション追加（resolveIdentifiers 前に設定が必要）
  - テスト: `tests/vba/CompileError.bas` の `Case_undefined_sub_call` / `Case_qualified_undeclared_obj`（TYPE: resolve）

- ✅ **識別子の大文字小文字混同の検出**: 同一モジュール/クラス/手続きスコープ内の変数・定数・型・クラス・プロシージャ・パラメーターを小文字化して比較し、ケースだけが異なる重複宣言をコンパイルエラーにする。Property Get/Let/Set の同名アクセサーは許可する | `identifier-case-collision.test.ts`
  - 例: `Dim assert As New Assert` — 変数 `assert` とクラス `Assert` は VBA では同一識別子
  - VBA はケースインセンシティブなため、コンパイル時に "識別子が重複しています" とすべきケース
  - 実装案: Lexer/Parser で識別子を正規化（小文字化）した後、同一スコープ内に同名の宣言があれば警告またはエラー

---

## リテラル型・演算子型の未実装項目

詳細は [`docs/internals/LITERALS_AND_OPERATORS.md`](../internals/LITERALS_AND_OPERATORS.md) を参照。

| 状態 | 項目 | 概要 |
|------|------|------|
| ✅ | **数値リテラルのサフィックス型情報保持** | `NumberLiteral` AST の `typeSuffix` フィールドで `TypeName(100&)` → `Long`、`TypeName(1.5!)` → `Single` など正しく返る | `typename.test.ts` |
| ✅ | **16進・8進リテラルの符号とサフィックス** | `&H8000`/`&HFFFF`/`&O177777` の符号付きビットパターン、`&H8000&` のLongサフィックス、基数リテラルの `TypeName`/`VarType` を実装 | `typename.test.ts` |
| ✅ | **LongLong/LongPtrリテラルの64ビット精度** | `^` サフィックス付き10進リテラルは元の桁列を保持し、`9223372036854775807^` などの境界値を丸めずに宣言型へ代入する | `identifier-type-suffix.test.ts` |
| ✅ | **サフィックス付きリテラルのオーバーフロー検出** | `100000%` → Overflow (Error 6)、`3000000000&` → Overflow (Error 6) | `typename.test.ts` |
| ✅ | **`1.0` リテラルの型推定誤り** | `isFloat` フラグにより `TypeName(1.0)` → `"Double"` を正しく返す | `typename.test.ts` |
| ✅ | **算術演算結果の型伝播** | `TypeName(1 + 1)` → `"Integer"`、`TypeName(6\2)` → `"Integer"`、`TypeName(6/2)` → `"Double"` など VBA 型昇格規則を AST レベルで実装（ランタイムラッパー不要） | `typename.test.ts` |
| ✅ | **Currency ランタイムラッパー実装** | `CCur()` は `VbaCurrency`（BigInt 固定小数点）を返す。`TypeName(c)` / `VarType(c)` は実値で判定。精度テスト参照 | `currency-bigint.test.ts` |
| ✅ | **Currency 算術演算の型保持** | `CCur(0.1)+CCur(0.2)` が厳密に `0.3`（float 誤差なし）。`Currency + Integer → Currency`、`Currency / Currency → Double` | `currency-bigint.test.ts` |
| ✅ | **Decimal 精度が IEEE 754 止まり** | `VbaDecimal` 内部を `mantissa: bigint` + `scale: number`（96-bit、scale 0-28）に刷新。28桁精度を実現 | `decimal-bigint.test.ts` |
| ✅ | **Decimal 算術演算の型保持なし** | `CDec(1) / CDec(3)` → VbaDecimal "0.3333333333333333333333333333"。+/-/*// も型保持 | `decimal-bigint.test.ts` |

---

## テストコードのリファクタリング残件

- ✅ **`tests/spec/` の `.evaluate(ast)` 形式を `evalVBASingle` / `evalVBAModules` に移行完了（33ファイル）**

---

## VBA エラー番号別の改善項目（Err.Number 対応）

現在のエンジンは多くの場面で `throw new Error(...)` （JavaScript の汎用エラー）を使っており、VBA の `Err.Number` に正しいエラー番号が設定されない。以下は Err.Number 別に整理した改善点。

凡例: `evaluator.ts:行番号` は該当 `throw new Error(...)` の場所を示す。

### Error 3 — Return without GoSub

- ✅ `evaluator.ts:3508` `Return without GoSub` → `throwVbaError(3, ...)` | `ControlFlowErrorTest.bas`

### Error 5 — Invalid procedure call or argument

- ✅ `evaluator.ts:202, 205` `VbaCollection.findIndex` でキー未発見 → `throwVbaError(5, ...)` | `CollectionErrorTest.bas`
- ✅ `evaluator.ts:2278` On…GoTo/GoSub のインデックス範囲外 → `throwVbaError(5, ...)` | `MiscErrorTest.bas`
- ✅ `Environment.set` 定数への代入 → VbaError 5 | `MiscErrorTest.bas`

### Error 9 — Subscript out of range

- ✅ `evaluator.ts:197` `VbaCollection.findIndex` の数値インデックス範囲外 → `throwVbaError(9, ...)` | `CollectionErrorTest.bas`
- ✅ `evaluator.ts:3614` `createMultiDimArray` で upper < lower → VbaError 9 | `MiscErrorTest.bas`
- ✅ `evaluator.ts:4189, 4296` 配列アクセス時に引数なし（インデックス省略） → `throwVbaError(9, ...)` | `MiscErrorTest.bas`
- ✅ 代入先が配列/辞書でない変数の添字アクセス (`x(0) = 5`) → `throwVbaError(9, ...)`
- ✅ 1D/2D 静的配列の読み取り・書き込みで `__vbaDimensions__` の lower/upper を超えたとき → `throwVbaError(9, ...)` | `array-subscript-bounds.test.ts`
- ✅ 次元数不一致（2D 配列を 1 インデックスでアクセス等） → `throwVbaError(9, ...)` | `array-subscript-bounds.test.ts`
- ✅ `ReDim` 後も更新された境界で正しくチェック | `array-subscript-bounds.test.ts`
- ✅ `Option Base 1` 時に index 0 アクセス → `throwVbaError(9, ...)` | `array-subscript-bounds.test.ts`

### Error 11 — Division by zero

- ✅ 除算演算子（`/`, `\`, `Mod`）で分母が 0 のとき → `throwVbaError(11, "Division by zero")` — 実装済み

### Error 13 — Type mismatch

- ✅ `evaluator.ts:2043` `For Each` に配列でもコレクションでもない値を渡したとき → `throwVbaError(13, ...)` | `ControlFlowErrorTest.bas`

### Error 35 — Sub or Function not defined

- ✅ `evaluator.ts:1646` `callProcedure` でプロシージャが見つからないとき → `throwVbaError(35, ...)` | `ProcNotFoundTest.bas`
- ✅ `evaluator.ts:4217` `evaluateCallExpression` で未知の識別子が引数付きで呼ばれたとき → `throwVbaError(35, ...)` | `ProcNotFoundTest.bas`
- ✅ `Environment.getProcedure` で同名プロシージャが複数モジュールに存在する（曖昧性エラー）→ VbaError 35

### Error 52 — Bad file name or number

- ✅ `evaluateWriteStatement`, `evaluateInputStatement`, `evaluateGetStatement`, `evaluateSeekStatement` でファイル未オープン → `throwVbaError(52, ...)` | `ControlFlowErrorTest.bas`
  - 注: `Print #`, `Line Input #`, `Put #` は既に `throwVbaError(52)` 実装済み

### Error 91 — Object variable or With block variable not set

- ✅ プロパティ代入先が `Nothing` / `undefined` のとき (Lines 2582, 2594, 2967) → `throwVbaError(91, ...)` | `ObjectErrorTest.bas`
- ✅ `evaluateCallExpression` MemberExpression ブランチで obj が `Nothing`/`null` → `throwVbaError(91, ...)` | `ObjectErrorTest.bas`
- ✅ `evaluateMemberExpression` で obj が `Nothing`/`null`/`undefined` → `throwVbaError(91, ...)` | `ObjectErrorTest.bas`

### Error 424 — Object required

- ✅ `evaluateSetStatement` で右辺が非オブジェクト (`Set x = 5`) → `throwVbaError(424, ...)` | `ObjectErrorTest.bas`
- ✅ プロパティ代入先がプリミティブのとき → `throwVbaError(424, ...)` | `ObjectErrorTest.bas`
- ✅ `evaluateMemberExpression` で obj がプリミティブ値のとき → `throwVbaError(424, ...)` | `ObjectErrorTest.bas`

### Error 429 — ActiveX component can't create object

- ✅ `evaluator.ts:2724` `New ClassName` でクラス未定義のとき → `throwVbaError(429, ...)` | `ProcNotFoundTest.bas`
- ✅ `evaluator.ts:3294` `CreateObject(progId)` で未サポートの ProgID のとき → `throwVbaError(429, ...)` | `ProcNotFoundTest.bas`

### Error 438 — Object doesn't support this property or method

- ✅ `evaluateCallExpression`: クラスインスタンスに存在しないメソッド → `throwVbaError(438, ...)` | `ObjectErrorTest.bas`
- ✅ `evaluateCallExpression`: JS オブジェクトに存在しないメソッド/プロパティ → `throwVbaError(438, ...)` | `ObjectErrorTest.bas`
- ✅ `evaluateMemberExpression`: プロパティが見つからない → `throwVbaError(438, ...)` | `ObjectErrorTest.bas`
- ✅ `evaluateDictionaryAccessExpression`: `!` アクセスが非 Dictionary → `throwVbaError(438, ...)` | `MiscErrorTest.bas`
- ✅ `evaluateCallExpression`: デフォルトプロパティが見つからない → `throwVbaError(438, ...)`
- ✅ `evaluateLetStatement`: デフォルト `Item` setter が見つからない → `throwVbaError(438, ...)`
- ✅ Dictionary アクセス時に引数なし → `throwVbaError(449, "Argument not optional")` に変更

### Error 449 — Argument not optional (Dictionary 引数なし)

- ✅ `dict()` / `dict` を引数なしで呼んだとき → `throwVbaError(449, ...)`

### Error 457 — This key is already associated with an element of this collection

- ✅ `evaluator.ts:179` `VbaCollection.add` でキー重複のとき → `throwVbaError(457, ...)` | `CollectionErrorTest.bas`

---

## 過去の仕様バグ履歴

サンプルコード検証・監査バッチなどで発見した過去のバグは、
[VBAエンジンバグ履歴](../bug-records/VBA_ENGINE_BUG_HISTORY.md)に分離して管理する。
