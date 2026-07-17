# VBA実装 TODOリスト（MS-VBAL仕様書準拠）

**進捗**: MS-VBAL 仕様書で個別セクションに名前が付いた **構文要素 / ステートメント / 式 / 標準ライブラリ関数 / 組み込みクラス** はすべて実装済みです。

ただし、これは「言語仕様書のリストに名前のある機能」の網羅率であって、**VBA のランタイム挙動（仕様書本文に書かれているが個別セクションを持たない暗黙の振る舞い）** までを 100% 満たしているわけではありません。実装が未確認のランタイム挙動については本ドキュメントの「[VBA ランタイム挙動](#vba-ランタイム挙動)」セクションを参照してください。

## 今後のロードマップ

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

凡例: ✅ 実装済み / ❌ 未実装 / ⚠️ 部分実装

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
| ⚠️ | P2 | VarPtr / StrPtr / ObjPtr | **制限実装**（ダミーポインター）。呼び出しごとに +4 増加する非ゼロ Long を返す（`_ptrCounter`、初期値 `0x10000`）。<br>**仕様**: `VarPtr(var)` は変数のメモリーアドレス、`StrPtr(str)` は BSTR バッファーのアドレス、`ObjPtr(obj)` は IUnknown ポインターを Long で返す（MS-VBAL 仕様外・Office VBA 非公式拡張）。<br>**制限**: 実アドレスを返せないため、同一変数への複数回呼び出しで値が変わる。`CopyMemory`・`RtlMoveMemory` 等の Win32 API 連携は機能しない。コードのパース・実行が通ることのみ保証。 | `varptr.test.ts` |

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
| ✅ | P2 | Put Statement | §5.4.5.11 | `filesystem-extra.test.ts` |
| ✅ | P2 | Get Statement | §5.4.5.12 | `filesystem-extra.test.ts` |

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

## 実装優先度ガイド

実用上の影響が大きい順に整理した優先実装候補：

### P0：最優先（基本制御フロー・頻出構文・標準関数）

1.  **Format / Format$** (§6.1.2.11.1.8) — 出力整形に多用
3.  **MsgBox / InputBox** (§6.1.2.11.1.x) — 対話型スクリプトの維持
4.  **Option Compare** (§5.2.1.1) — 文字列比較の正確性
5.  **Attributes** (§5.2.3.1.6) — ソースエクスポートファイルの直接読込
6.  **CreateObject (拡張)** (§6.1.2.8.1.4) — 外部オブジェクト連携の強化
7.  **エラーハンドリング** (§5.4.4) — `On Error GoTo`, `Resume` の完全な準拠

### P1：重要（標準ライブラリの拡充）

7.  **文字列関数**: `Filter`, `StrConv`
8.  **変換関数**: `CByte`, `CCur`, `CDec`, `CSng`
9.  **数学関数**: `Abs` 等の未実装エッジケース
10. **制御文**: `Assert` ステートメント

### P2：低優先度（特殊用途・レガシー）

11. 財務関数 (Financial Module) 全般
12. ファイル操作ステートメント (Open/Print等) — Node.js環境での設計が必要
13. レガシー構文: `GoSub` 等

## テストダブル（スタブ・モック・仮想化）の実装予定

VBA Runner の性質上、テストの安定性および検証精度の向上のために以下の実装を今後検討します。

- **MockWorksheet の改善**:
    - ✅ **1D 配列での書き込み**: VBA では `Range("A1:A4").Value = Array(10,20,30,40)` のように 1D 配列を代入すると、配列を1行分の列値として解釈し全行に繰り返し適用する（例: A1〜A4 すべてが 10）。`MockWorksheet.setCellValue` / `Range().Value =` の両方で対応。 | `mock-worksheet-address.test.ts`
    - ❌ **配列サイズ不一致時のエラー検出**: 範囲と配列のサイズが合わない場合、VBA は Error 1004 を発生させるが、MockWorksheet はサイレントにスキップ/空文字填充する。
    - ✅ **`Range().Value =` の書き戻し**: VBA では `ws.Range("A1:B3").Value = array` でセルに書き込めるが、MockWorksheet では `setCellValue` を使わないとセルに反映されない。 | `mock-worksheet-address.test.ts`
- **`VBARunner.setConstants()` で注入した値の上書き制限**:
    - ✅ 現状、`setConstants()` は通常の変数として値を設定するため、VBA コード内で `xlUp = 999` のように上書きできてしまう。`Evaluator.setConstant()` を追加し `VBARunner.setConstants()` から呼ぶよう変更。VBA コード側からの代入は既存の `Environment.setConstant()` / `isConstant()` 機構により Error 5 になる。 | `set-constants-protection.test.ts`
- **副作用の検証機能 (Spy / Verify)**:
    - `Shell` や `MsgBox` 等のスタブ・モック関数に対して渡された引数（実行コマンド名やメッセージ内容）を、テストコード側からプログラム的に検証（アサーション）できる仕組み。
    - 現状：`Debug.Print` 相当のコンソール出力のみ。
- **日付・時刻の仮想化 (`Date`, `Now`, `Time`)**:
    - 現状：システムの現在日時を返却。
    - 検討内容：テスト実行時に特定の固定日時を「現在時刻」として設定・返却できる仕組み（テストの再現性向上のため）。
    - 優先度：未定。
- **仮想レジストリ**:
    - ✅ `GetSetting` / `SaveSetting` / `DeleteSetting` / `GetAllSettings` をメモリ上のマップでエミュレート。
- **外部COMオブジェクトのスタブ・フェイク拡充**:
    - `MSXML2.XMLHTTP` 等の頻出オブジェクトのメソッド・プロパティの実装。

## vba-analyzer の既知の制限

### 即値引数検出（`Range`/`Cells`/`Sheets`/`Worksheets`）

- ✅ **Range 変数経由のアクセス**: `rng(3, 5)` / `rng.Item(3, 5)` のように `Dim rng As Range` と宣言された変数を介したアクセスを検出。`findMagicLiteralsInCalls` 内で Dim 宣言を先行走査して型マップを構築し、visit 時に照合する軽量実装で対応（CFG/reaching-defs は不使用）。なお既存の CFG + reaching-defs + live-vars インフラ（`def-use-analyzer.ts`）を活用すれば Dead Store 検出・未初期化変数の精密検出・Const 候補検出なども将来実装可能。 | `magic-literals.test.ts`

## 仮想ファイルシステム (VFS) の実装状況と課題

Webブラウザおよびテスト環境向けの仮想ファイルシステム (`MemoryFileSystem`) の実装状況です。

### 機能マトリックス

| 状態 | 機能カテゴリ | 項目 | 説明 |
| :---: | :--- | :--- | :--- |
| ✅ | **基本 I/O** | `readSync` / `writeSync` | バッファー/文字列の同期読み書き。ハンドル位置の自動更新対応。 |
| ✅ | | `openSync` / `closeSync` | ファイルオープン。`w` フラグ時の truncation (切り詰め)、`a` フラグ時の末尾追記（2026-06-25 修正、後述）対応済み。 |
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
| ❌ | | **排他制御** | `Lock` / `Unlock` ステートメントのエミュレーション。 |
| ⚠️ | **互換性** | **バイナリ/テキスト** | `Binary` / `Random` / `Input` / `Output` 各モードの厳密な挙動。 |

### VFS 開発ロードマップ (TODO)

#### 1. コア機能の強化
- ✅ **仕様バグ修正（2026-06-25）: `Open ... For Append` が追記ではなく上書きになる（サイレントなデータ消失）**
  - 原因: `MemoryFileSystem.openSync()` がファイルハンドルの書き込み開始位置 `pos` を `flags` に関係なく常に `0` で初期化していたため、`writeSync` が `Append` モードでも常にファイル先頭から上書きしていた。2回目以降の `Open ... For Append` のたびに、それまでの追記内容が消える
  - 症状: 例外は発生せず、ログファイル追記のような典型的な用途でデータが静かに失われる。`NodeFileSystem`（実ファイル版、`vba-run` CLI 等で使用）は Node 標準の `fs.openSync` に委譲しているため影響を受けず、`VBARunner` が使う `MemoryFileSystem` のみの問題
  - 修正: `flags === 'a'` の場合、既存ファイルが存在すればその長さを `pos` の初期値にするよう変更 | `filesystem.test.ts`
- ✅ `MemoryFileSystem` にワイルドカードマッチングロジックの実装（`Kill` への統合）。 | `wildcard-kill-dir.test.ts`
- ✅ `MemoryFileSystem` 内部での `cwd` (Current Working Directory) の保持と `ChDir` 対応。 | `chdir-curdir.test.ts`
- ❌ `statSync` における VBA 属性（Read-only, Hidden 等）のシミュレーション。

#### 2. ブラウザ環境最適化
- ❌ `FileSystem` インターフェースの非同期版 (`read`, `write` 等) の検討（ブラウザのメインスレッドをブロックしないため）。
- N/A `IndexedDB` をバックエンドとした `PersistentFileSystem` の実装。（実装予定なし：本エンジンの用途範囲外）

#### 3. テスト環境の改善
- ✅ `filesystem-extra.test.ts` 等のテストコードを、Node.js `fs` 直接参照から `this.fs` (抽象インターフェース) 参照へリファクタリング。（`MemoryFileSystem` + `Evaluator({ fs: vfs })` で対応済み）
- ❌ 複数の `Evaluator` インスタンス間で単一の `MemoryFileSystem` を共有するためのテスト用設定の追加。

#### 4. バイナリ操作の極致
- ❌ `Put` / `Get` における固定長文字列、UDT、多次元配列のバイナリレイアウトの完全な VBA 互換性検証。

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
| ⚠️ | §5.4.2.1 | `Call .Method arg`（parens なし・With 式 ImplicitWithObject 形式）が CallExpression でないためエラー | `Call .Draw 10, 20` |
| ⚠️ | §5.4.2.9 | `If x > 0 Then 100`（行番号ジャンプ）が GoTo として機能しない | レガシー BASIC スタイル |
| ⚠️ | §5.4.2.10/§5.6.9.5 | `Case Is ><5` / `><` `=<` `=>` 代替比較演算子（レキサーが別トークンに分割） | `Case Is ><5` |
| ⚠️ | §5.4.2.12/14 | `Go To` / `Go Sub`（2語形式）→ `Go` が識別子扱いでパースエラー | レガシー |
| ⚠️ | §5.4.3.1 | `Dim Shared x As Integer`（VBA6 方言） | レガシー |
| ⚠️ | §5.4.5.1 | `Open "f.txt" For Random Access Read Shared As #1`（`Shared` スタンドアロン） | ネットワークファイル |
| ⚠️ | §5.4.5.4/5 | `Lock #1, To 100`（start 省略の `To end` 形式） | 稀な書き方 |
| ⚠️ | §5.4.5.6 | `Line Input #1, arr(i)`（variable が配列要素） | 稀な書き方 |
| ⚠️ | §5.2.3.3 | UDT メンバー名として `End` が使用不可（ループ終端キーワードと衝突） | `Type T: End As Long: End Type` |
| ⚠️ | §5.2.4.1.2 | `Implements SomeLib.IInterface`（ドット修飾名）→ 1 トークンしか読まずパースエラー | `Implements ADODB.ICommand` |
| ⚠️ | §5.3.1.3 | `Function Foo$()` 等 TYPED-NAME サフィックス付き宣言名 → `Foo()` 呼び出しと不一致 | レガシー |
| ⚠️ | §5.3.1.5 | `ByVal Optional x As Integer`（ByVal が先）→ `Optional` をパラメーター名と誤認しパースエラー | 稀な書き方 |
| ⚠️ | §5.6.5 | `1.5D10`（VBA Double `D` 指数表記）→ レキサーが `D`/`d` を未対応（§3.x レキサー問題） | 科学技術計算 |
| ⚠️ | §5.6.9.5 | 連鎖比較演算子の優先度誤り（`a = b < c` → `a = (b < c)` と解析、正しくは `(a = b) < c`） | 極めて稀 |
| ⚠️ | §5.6.13.1 | `Foo(ByVal x)`（呼び出し側 ByVal）→ 呼び出し引数で `ByVal` が未処理でパースエラー | 古い COM スタイル |
| ⚠️ | §5.6.15 | `With dict: !Key = v`（行頭 `!` を With 辞書アクセスとして未処理）→ `!` 単体でエラー | With ブロック内 `!Key` |
| ⚠️ | §5.6.16.7 | `Dim x As A.B.C`（type-expression の 2 段階以上ドット修飾）未対応 | 稀な書き方 |
| ⚠️ | §5.6.16.8 | `AddressOf NS.Module.Proc`（多段階ドット）→ 1 段階のみ対応 | 稀な書き方 |

---

## VBA ランタイム挙動

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
- ⚠️ **WithEvents 変数の生存期間**: 親オブジェクト破棄時のイベントハンドラー解除 (制限事項: RaiseEventコア存在; ハンドラー自動登録・検出機構未実装、イベントパラメーター解析未対応) | `withevents-lifecycle.test.ts`
  - 実装しても**同期呼び出しのみ対応**（`src.Fire` → `src_TestEvent` が即時実行される形）
  - `DoEvents` + `Sleep` によるイベント待ちループ（wait/notify パターン）は不可。evaluator がシングルスレッド同期のため、ループ中に外から RaiseEvent を発火する主体が存在しない
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

### エラーハンドリング

- ✅ **仕様バグ修正（2026-06-26）: 組み込み実行時エラーの `Err.Description` が `"Run-time error 'N': ... (line X)"` の枠組みテキストを含む**
  - 原因: `throwVbaError()` が TypeScript 側に投げる例外の `.message` 用に枠組み付き文字列を生成しており、`executeStatements` の catch ブロックがその `e.message` をそのまま `errObj.description` に転記していた。`Err.Raise`（`ErrObject.raise()`）経由のカスタムエラーは元々生のメッセージを直接 `description` に入れていたため問題なかったが、組み込みエラー（Type Mismatch・Division by Zero 等）だけ枠組み付きになっていた
  - 修正: `throwVbaError()` が投げる例外に `vbaBareMessage`（枠組みなしの生のメッセージ）を追加で持たせ、`errObj.description` への転記時はそちらを優先する。TypeScript 側に渡る `.message` 自体（`eval()`/`run()` の catch で見えるエラーメッセージ）は枠組み付きのまま変更していない | `err-object-full.test.ts`
- ✅ **仕様バグ修正（2026-06-26）: `vbObjectError` 定数が未実装**
  - 症状: カスタムエラー番号を生成する MS 推奨の慣用句 `Err.Raise vbObjectError + n` が `Option Explicit` 配下で未宣言変数エラーになる
  - 修正: `vbobjecterror = -2147221504`（§6.1.2.10）を組み込み定数として追加 | `err-object-full.test.ts`
- ✅ **仕様バグ修正（2026-06-25）: `On Error Resume Next:` / `On Error GoTo Label:` の直後に `:` で後続文を連結すると消える**
  - 原因: `parser.ts` の `parseOnErrorStatement()` が "Resume Next" 部分を読み取るループで `Newline`/`EOF` のみを終端条件にしており、`:`（`OperatorColon`）で止まらなかったため、同一行の `:` 以降の文すべてが label 文字列として読み込まれ AST から消えていた
  - 修正: 既存の `isAtTerminator()`（`Newline`/`EOF`/`OperatorColon` を終端とする）を使うよう変更 | `on-error-colon.test.ts`
- ✅ **仕様バグ修正（2026-06-25）: `VBARunner.eval()` 内で `On Error Resume Next`/`On Error GoTo` が機能しない**
  - 原因: `evalExpression()` の複数文フォールバックが `evaluateModule()` の単純な逐次実行ループを使っており、`On Error`（`Resume Next`/`GoTo`/`Resume`）の分岐ロジックを持つ `executeStatements()`（手続き本体実行や `run()` 経由では使われる）を経由していなかった
  - 修正: `evalExpression()` のフォールバックを `executeStatements(program.body, 0)` に変更。さらに `eval()` 呼び出しごとに `errorHandlerLabel`/`errorHandlingMode`/`isInErrorHandler`/`lastErrorIndex`（および `currentSourceModule`）を退避・復元し、`eval()` を独立したトップレベル呼び出しフレームとして扱うよう統一 | `eval-on-error.test.ts`, `eval-after-load.test.ts`
- ✅ **仕様バグ修正（2026-06-25）: クラスメソッド呼び出しが On Error 状態を呼び出し元に漏らす**
  - 原因: `callClassMethod()` が `errorHandlerLabel`/`errorHandlingMode`/`isInErrorHandler`/`lastErrorIndex` を退避・復元していなかった（通常の手続き呼び出し `execProcBody` は退避・復元している）。クラスメソッド内で `On Error` ハンドラーに入った状態（`isInErrorHandler=true`）のまま例外が呼び出し元に伝播すると、呼び出し元の `executeStatements` がこれを引き継いでしまい、呼び出し元自身の `On Error GoTo`/`Resume Next` への分岐が機能せずエラーがそのまま素通りしていた
  - 修正: `execProcBody` と同様に4フィールドを退避・リセット・復元するよう `callClassMethod` を修正 | `class-method-on-error-isolation.test.ts`
- ✅ **仕様バグ修正（2026-06-25）: `VBARunner.eval()` で定義したプロシージャが `Option Explicit` の静的検査対象にならない**
  - 原因: `resolveIdentifiers()`（Pass 2）は `VBARunner` 構築時にロードしたモジュール群のみを `checkOptionExplicit` の対象にしており、構築後に `eval()` で定義したプロシージャは一度も解析されず `optionExplicitViolations` に登録されなかった
  - 修正: `evalExpression()` の複数文フォールバックで都度 `checkOptionExplicit` を実行し、違反があれば `optionExplicitViolations` に追加登録するよう修正（既存の登録は維持） | `eval-option-explicit.test.ts`
- ✅ **`On Error Resume Next` 下の暗黙の `Err.Clear`**: 正常文実行時のクリア（または非クリア）タイミング | `err-clear-timing.test.ts`
- ✅ **`Resume` の対象決定**: エラー発生点・Resume Next・Resume <label> の正確な制御フロー | `resume-statement-target.test.ts`
- ✅ **エラーハンドラー内での再帰的なエラー発生**: スタックフレームのリセット規則 | `recursive-error-handling.test.ts`
- ⚠️ **`Erl` 関数**: エラー発生行番号の取得（制限事項: MS-VBAL spec未定義、line number tracking未実装）★優先度低: VBA ソース側に数値行ラベル（`10 If ...` 形式）がないと Erl は常に 0 を返すため、実用コードでの使用頻度が低い | `erl-function.test.ts`

### プロシージャ呼び出しの細部

- ✅ **モジュール修飾付きプロシージャ呼び出し**: `ModuleName.ProcedureName()` による複数ファイル間での同名関数の区別
  - 実装: `Environment.setProcedureWithModule()` / `getProcedureFromModule()` でモジュール名付き登録・検索
  - VBA 名前解決優先度（修飾なし呼び出し）:
    1. グローバルスコープ（モジュール名なし）← 最優先
    2. 単一モジュール内のプロシージャ（曖昧でない）
    3. 複数モジュール → 曖昧性エラー（修飾必須）
  - 曖昧性検出: 複数モジュールに同名プロシージャがある場合、実行時に詳細エラーを発生
  - テスト: `module-qualified-calls.test.ts`, `ambiguous-procedure-call.test.ts`, `evaluator-scoping.test.ts`
- ✅ **Fix: モジュール修飾呼び出しが呼び出し先のランタイムエラーを握りつぶす** | `runtime-error-trace.test.ts`
  - 原因: `evaluateCallExpression` の `Module.Proc` 検出が `try { callProcedure } catch { フォールスルー }` で**あらゆる例外**を飲み込み、member access に落ちていた
  - 症状: `ModB.DoWork` の本体がゼロ除算（Error 11）を投げても握りつぶされ、未定義オブジェクト `ModB` への member access として Error 91 にすり替わる
  - 修正: `getProcedureFromModule` で**事前に存在確認**し、見つかった場合のみ `callProcedure` を呼ぶ（try/catch を廃止）。呼び出し先のエラーは正しく伝播する
- ✅ **Fix: `VBA.Func()` が標準ライブラリ関数を呼ぶ強制力を持たない** | `vba-stdlib-qualifier.test.ts`
  - VBA 仕様: `VBA.InStr(...)` のような型ライブラリ修飾呼び出しは、ユーザーが同名の関数を定義していても必ず標準ライブラリを呼ぶ
  - 原因: `evaluateCallExpression` の `VBA.X` 処理が callee を非修飾 `X` に差し替えて再帰していたため、`getProcedure`（ユーザー定義優先）の経路に乗っていた
  - 修正: `env.getConst(name)` で `variables`（組み込み関数テーブル）のみを検索し、`procedures`（ユーザー定義）をスキップして直接呼び出す
- ✅ **Fix: `VarType(VBA)` / `VarType(Module1)` のようにプロジェクト名・モジュール名を値として使った場合にエラーにならない** | `namespace-as-value-error.test.ts`
  - VBA 仕様: `VarType(VBA)` → コンパイルエラー「プロジェクトではなく、変数またはプロシージャを指定してください」。モジュール名も同様。
  - 修正: `vba-types.ts` に `VbaNamespaceRef` センチネルクラスを追加し、`VBA` を `kind='project'` で、各モジュール名を `kind='module'` で env に事前登録。`Identifier` 評価時に `VbaNamespaceRef` を検出してエラーを投げる。修飾形式（`VBA.X`、`Module1.Proc`）は早期リターンで正常動作を維持。
  - **MS-VBAL 上の「プロジェクト」の定義・実装詳細**: `REFERENCE.md` §「外部ライブラリの参照設定と `New ProjectName.ClassName`」を参照。`registerExternalObject('X.Y', factory)` でプロジェクト名 X が `VbaNamespaceRef kind='project'` として自動登録される仕様を含む。
- ✅ **モジュール修飾付き変数/定数アクセス**: `Module1.A` 形式でモジュールレベルの変数・定数を参照
  - 実装: Const は module-qualified キー (`module1:a`) で格納（不変なので複製コピーで同名競合も区別可）。変数は `moduleVarRegistry` に登録し参照時は非修飾名で引く
  - `evaluateMemberExpression` でオブジェクト評価前に台帳チェック（`Environment.get` の暗黙ゼロ初期化による誤検知を防止）
  - 制限事項: 同名のモジュールレベル変数（Const でない）が複数モジュールに存在する場合、最後の代入が勝つ（区別不可）
  - テスト: `module-qualified-access.test.ts`
- ✅ **ByRef での文字列・配列・オブジェクトの参照保持**: 文字列・数値・Boolean・配列・オブジェクト全て正常動作 | `byref-reference-preservation.test.ts`
- ✅ **仕様バグ修正（2026-06-25）: `VBARunner.run(name, args)` で ByRef パラメーターに代入しても呼び出し元の `args` 配列に反映されない**
  - 「あるべき姿」: VBA の既定の引数渡しは ByRef（明示的な `ByVal` がない限り）。`Sub Increment(ByRef n As Long): n = n + 1` を `run('Increment', [5])` で呼んだ後、呼び出し元が渡した `args` 配列は `[6]` になっているべき（「ステータスコード + メッセージを ByRef の out パラメーターで返す」という VBA で一般的なパターンを TypeScript 側で受け取れる必要がある）
  - 「修正前の現状」: `callProcedure()` は引数を `localEnv` に値としてコピーするだけで、JS の数値・文字列・Boolean はプリミティブで参照を共有しないため、Sub 内で ByRef パラメーターに代入してもエラーは出ないが `args` 配列には一切反映されなかった（`evaluateCallExpression` 経由の VBA-to-VBA 呼び出しは `originalExpr` を使った別の書き戻し機構を既に持っており問題なし。`callProcedure()` 経由＝主に `VBARunner.run()` からの直接呼び出しのみが対象）
  - 修正: `callProcedure()` で `execProcBody()` 呼び出し後、`ParamArray` でも `ByVal` でもないパラメーターについて `localEnv` から最終値を読み、呼び出し元の `args[i]` に書き戻すよう変更 | `run-byref-writeback.test.ts`
- ✅ **ParamArray の境界ケース**: 0 個渡し、配列を 1 つだけ渡したときの展開規則、ByRef semantics (spec §5.3.1.5) | `paramarray-edge-cases.test.ts`
- ✅ **Optional パラメーターの IsMissing 判定**: デフォルト値ありと未指定の区別
- ✅ **Property Get/Let/Set の解決順序**: 同名で混在した場合の優先度 | `property-resolution-order.test.ts`
- ✅ **暗黙の Let（`Call` なしの呼び出し）**: 戻り値が破棄される / されない場面 | `implicit-let.test.ts`
- ✅ **引数の数の検証**: プロシージャ呼び出し時に引数の個数が定義と異なる場合にエラーを発生させる (Error 450 / 449) | テスト: `ArgCountTest.bas`
  - 制限事項: VBA 本来はコンパイルエラーだが、本インタープリターでは実行時エラーとして発生させる
- ✅ **位置引数の省略（`Func(a,,c)`）**: `,,` で中間引数を省略した場合に Optional のデフォルト値が適用される | テスト: `missing-arg.test.ts`
  - 修正: 省略スロット（`MissingArgument` ノード）が `positionalArgs` に含まれていてもデフォルト値分岐に落ちるよう修正
  - 仕様バグ修正（2026-06-21）: 省略スロットが Optional でも defaultValue もない必須パラメーターの位置に来た場合（`Foo(1, , 3)` で `b` が必須）、従来は引数個数チェック（`checkArgCountGeneric` は占有スロット数しか見ない）をすり抜けてサイレントに `0`/`Empty` で実行されていた。VBA 本来はコンパイルエラーになるべきケースのため、`checkNoGapOnRequiredParam(params, argExprs)` を新設し 449 (Argument not optional) を投げるよう修正。標準モジュールの Sub/Function（インライン束縛）とクラスメソッド呼び出し（`callClassMethod` 呼び出し前の4箇所）の両方に適用 | テスト: `missing-arg.test.ts`

### 数値・日付の演算

- ✅ **`\`（整数除算）の VBA 厳密仕様**: 引数の自動型変換（Variant → Integer/Long）と余り処理
- ✅ **`Mod` 演算子のオーバーロード**: 浮動小数 vs 整数の挙動差
- ✅ **`#1/1/2024#` 型 Date リテラルの月日順序**: ロケールに依存しない解釈の保証 | `date-literal-parsing.test.ts`
- ✅ **`DateAdd("m", 1, ...)` 等の月末ロールオーバー規則**: Jan31 + 1 month = Feb28/29 | `dateadd-monthend-rollover.test.ts`

### 文字列処理

- ✅ **仕様バグ修正（2026-06-25）: `Join()` が下限付き固定配列（`Dim a(1 To n)`）の先頭に余分な空要素を混入させる**
  - 原因: `Dim a(1 To n)` の物理ストレージは `LBound` 分の隠し添字（添字0など）を含む JS 配列で確保され、`UBound`/`LBound`/添字境界チェックは `vbaBase`/`__vbaDimensions__` を見て隠しているが、`join` ビルトインは `arr.join(del)` を物理配列にそのまま適用していた
  - 修正: `vbaBase`（既定 0）でスライスしてから `join` するよう修正 | `split-join.test.ts`
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

- ✅ **仕様バグ修正（2026-06-26）: `eval()` の高速パスが式実行中の本物のランタイムエラーを握りつぶし、無関係な構文エラーに化けさせる**
  - 原因: `evalExpression()` は「単一式としてパースできるか」を try/catch で試すが、その catch ブロックが `parseExpressionPublic()` の呼び出しだけでなく後続の実行（`callProcedure`/`evaluateExpression`）まで囲んでいた。式の中で呼んだ関数が `Error 91` 等の本物のランタイムエラーを投げても「パース失敗」として握りつぶされ、文として再解析した際に得られる無関係な構文エラー（または別の挙動）に置き換わっていた
  - 症状: `eval('Foo()')`（`Foo` が内部で実行時エラーを投げる）が `run('Foo', [])` と異なる、誤解を招くエラー（または `Foo()` を裸の statement として再解析したことによる "syntax error"）になる。デバッグ時に本当の原因が一切見えなくなる
  - 修正: パース段階（`parseExpressionPublic()` 呼び出しと `fullyConsumed`/`isStatementAmbiguous` 判定）のみを try/catch で囲み、実行（`callProcedure`/`evaluateExpression`）はその外側で行うよう構造を変更 | `eval-real-error-propagation.test.ts`
- ✅ **仕様バグ修正（2026-06-25）: `eval()` に複数行コードを渡すと1行目だけが式として誤評価され、2行目以降が黒く無視される**
  - 原因: `evalExpression()` の「単一式として全体を消費したか」判定が `parseExpressionPublic()` 直後のトークンが `Newline` であることだけを見ており、その先に何があるか確認していなかった。VBA の `=` は代入文とも等価比較式とも解釈できるため、`"x = 10\nDebug.Print 1"` のような入力では1行目 `x = 10` が単独で「等価比較式」としてパースしきれてしまい、直後が改行であることから「全体を消費した」と誤判定していた
  - 症状: `x = 10` が代入ではなく比較として評価され（`x` は未代入のまま）、2行目以降（`Debug.Print 1` 等）は一切実行されない。`eval()` に Sub/Function 定義を含まない複数行スニペットを渡す、ごく自然な使い方で発生する
  - 修正: 直後が `Newline` の場合は、それ以降（複数の `Newline` を挟んでも）が `EOF` であることまで確認するよう変更。単一式＋末尾改行のみという既存の正しいケースは維持される | `eval-multiline-truncation.test.ts`
- ✅ **仕様バグ修正（2026-06-25）: `eval()` に裸の代入文（`x = 10` 単体、`arr(1) = "a"` 単体）を渡すと代入が一切実行されない**
  - 原因: 上記と同根。`parseExpressionPublic()` が `x = 10` を「`x` と `10` を比較する `BinaryExpression`」として入力全体ぴったり消費する形でパースしてしまい、高速パスの「単一式として全体を消費した」判定を満たしてしまっていた。代入は一度も実行されず、比較結果の真偽値だけが返っていた
  - 症状: `eval('x = 10')` の直後に `eval('x')` しても未代入のまま。配列要素も同様（`eval('arr(1) = "a"')` の代入が反映されない）。`Dim` と代入を別々の `eval()` 呼び出しに分けるという自然な使い方で発生する
  - 修正: 入力全体を消費した式がトップレベルの `=` を持つ `BinaryExpression` の場合は高速パスを使わず、文として解析・実行するフォールバックに委ねるよう変更（実 VBA でも裸の `lhs = rhs` 文は常に代入と解釈されるため、副作用なしの比較式として読まれる余地はない） | `eval-bare-assignment.test.ts`
- ✅ **仕様バグ修正（2026-06-25）: `eval('x + 1')`（単独）と `eval('... : x + 1')`（複数文の最後）の結果が一致しない**
  - 前提（実 VBA 仕様の確認）: VBA の statement 文法では、識別子で始まり `+`/`-`（および `=`）が続く裸の文は、単独行でも複数文中でも**常に同一の解釈**になる。一方 `<`/`&`/`And` 等の比較・連結・論理演算子は statement としては成立せず Parse error になるため、この曖昧性自体が存在しない（`x < 10` を文として解析するとパースエラー）。検証結果:
    | 入力（文として解析） | 解析結果 |
    |---|---|
    | `x - 1` | `CallStatement`（`x` を引数 `-1` で呼び出す） |
    | `x + 1 + 2` / `x + 1 - 2` | `CallStatement`（深い `+`/`-` チェーンも同様に1つの呼び出しに集約） |
    | `Foo(1) + 1` | `CallStatement`（明示括弧付き呼び出しが左辺の場合。後述の追加修正で式評価に変更） |
    | `x.Bar + 1` | `CallStatement`（メンバアクセスが左辺でも同様） |
    | `1 + x`（識別子が左辺でない） | `LabelStatement` + 別文（数値リテラル開始は行ラベルと解釈され、別カテゴリの曖昧性のため対象外） |
    | `(x) + 1`（括弧で囲む） | Parse error（括弧で明示的に式扱いになり曖昧性が解消される） |
    | `x < 10` / `x & "a"` / `x And y` | Parse error（statement として成立しないため曖昧性なし。式としての評価が唯一の解釈） |
  - 「あるべき姿」と「修正前の現状」:
    | x の種類 | 文脈 | あるべき結果（実VBA仕様） | 修正前の現状 |
    |---|---|---|---|
    | 変数（例 `x=5`） | 単独 `eval('x + 1')` | `x` は呼び出し不可能なため `Error 424` | ❌ `6`（算術加算として誤評価） |
    | 変数 | 複数文中 `... : x + 1` | `Error 424` | ✅ `Error 424`（既に正しい） |
    | Function（必須引数1個） | 単独 | 引数 `1` で呼ばれ、戻り値は文として捨てられる（エラーなし） | ❌ `1`（`x` が暗黙的に `Empty`(0) 扱いされ `0+1` を返す。呼び出しは発生しない） |
    | Function（必須引数1個） | 複数文中 | 同上 | ✅ 同上（既に正しい） |
    | Sub（必須引数1個） | 単独 | 引数 `1` で呼ばれる（エラーなし） | ❌ `1`（同上、呼び出しは発生しない） |
    | Sub（必須引数1個） | 複数文中 | 同上 | ✅ 同上（既に正しい） |
  - 原因: 上記2件と同根。`evalExpression()` の高速パスは「トップレベル演算子が `=` か」しか見ておらず、`+`/`-` で識別子が左辺（再帰的に、ネストした `+`/`-` チェーンの最も左の葉まで）に来るケースを見ていなかった
  - 修正: 高速パスの判定に、トップレベルが `+`/`-` の `BinaryExpression` で、左辺を `+`/`-` の連鎖を通して再帰的に辿った最終的な葉が `Identifier`/`MemberExpression`（= 呼び出し可能な形）である場合も「statement として解析すべき」に含めるよう拡張。括弧で明示的に囲まれた場合（`(x) + 1`）は曖昧性が解消されるため対象外とする | `eval-statement-ambiguity.test.ts`
  - **既知のスコープ外（意図的に対応しない）**: `eval('Foo(1, 2)')` のように明示括弧付きの関数呼び出しが**単独で** `eval()` に渡された場合、現状は戻り値をそのまま返す（複数文中で同じ呼び出しが行われた場合は statement として戻り値を捨てる、という非対称性が残る）。これは「関数呼び出しの結果をそのまま取得する」という最も基本的で有用な `eval()` の使い方を壊さないよう、意図的に対象外としている
- ✅ **仕様バグ修正（2026-06-27）: `eval()` で `UBound(arr) + 1` / `Len("x") + 1` 等の組み込み関数戻り値への算術演算が `Error 424: Object required` になる**
  - 原因: `isCallableLeftmostLeaf()` が `CallExpression` ノードに対して `true` を返していたため、`UBound(arr) + 1` の左辺（`CallExpression`）が「曖昧な呼び出し文」と誤判定された。fast-path（式評価）がスキップされ、フォールバックの文パーサーが `UBound(arr + 1)`（配列オブジェクトに 1 を加算）として誤解析し、Error 424 が発生していた。括弧ワークアラウンド `(UBound(arr)) + 1` では外括弧でルートノードが `ParenthesizedExpression` になるため fast-path が正しく使われていた
  - 修正: `isCallableLeftmostLeaf()` の `CallExpression` ケースを `return false` に変更。明示括弧付きの関数呼び出し `foo(args)` は括弧で引数リストが確定しており、`foo args+1` のような暗黙 Call 文との曖昧性はない | `eval-builtin-arithmetic.test.ts`
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

### Parser バグ修正

- ✅ **仕様バグ修正（2026-06-25）: `Function Foo() As String()`（配列戻り値型）の `()` が読み捨てられる** | `function-array-return-type.test.ts`
  - 原因: 戻り値型パース（`As <Type>`）が型名トークンのみを読み、後続の `()`（配列マーカー）を見ていなかったため `returnType` が単なる `"String"` になり、`()` トークン自体は本体パース側で空文として黙って消費されていた
  - 症状: `evaluator.ts` の戻り値代入処理が戻り値変数をスカラー String 型と誤認し、`coerceToDeclaredType` 経由で配列に `CStr()` を適用 → `IsArray(GetArr())` が `False` になり、配列要素アクセスで Error 424
  - 修正: `()` を検出したら `ProcedureDeclaration.returnsArray` フラグを立て、`evaluator.ts` 側は `returnsArray` のときスカラー型 coercion を適用しないよう修正（`execProcBody`/`callClassMethod` 両方）

- ✅ **仕様バグ修正（2026-06-26）: クラスフィールドの固定長配列が Empty のまま初期化されず、外部・内部アクセスともに Error 9/438 になる** | `class-module.test.ts`
  - 原因1: `createInstanceFromDef` がフィールドをループ初期化する際に `decl.isArray` チェックがなく、固定長配列が `vbaEmpty` のまま `instanceEnv` に格納されていた。ローカル変数の `evaluateVariableDeclaration` では `decl.isArray → createMultiDimArray()` が正しく呼ばれていた
  - 原因2: 読み取り側の `evaluateCallExpression`（VBA クラス分岐）でも、プロシージャが見つからない場合に配列フィールドへのインデックスアクセスが未実装だった（Error 438）
  - 原因3: 書き込み側の `evaluateAssignmentToVariable`（`obj.Field(i) = val` 形式）でも同様に配列フィールドへの代入が未実装だった
  - 修正: ① `createInstanceFromDef` に `isArray` 分岐を追加（`evaluateVariableDeclaration` と同等の処理）、② 読み取りパスに配列フィールドのインデックス解決を追加、③ 書き込みパスに同様の処理を追加

- ✅ **仕様バグ修正（2026-06-26）: `run()`/`eval()` の戻り値 `Nothing` が JS Symbol のままで、テンプレートリテラルで TypeError / if 判定が誤る** | `test-libs/test-runner.ts`
  - 原因: `vbaNothing = Symbol('vbaNothing')` が `VBARunner.run()`/`eval()` の公開 API 境界で変換されていなかった。Symbol はテンプレートリテラルへの暗黙変換が禁止されており、オブジェクト同様に truthy なため `if (!result)` での Nothing 検出も誤動作した
  - 修正: `VbaBoolean → boolean` 変換と同様に、戻り値が `vbaNothing` のとき `null` に変換するよう `run()`/`eval()` に追加

- ✅ **仕様バグ修正（2026-06-26）: `Dim w As New T` を一度もアクセスせず `Set w = Nothing` すると `Class_Initialize`/`Class_Terminate` が呼ばれない** | `circular-reference-terminate.test.ts`
  - 原因: vba-runner は `Dim w As New T` を遅延インスタンス化（初回アクセス時に生成）で実装していた。`Set w = Nothing` はオブジェクトへの参照とみなされ、実 VBA ではここでインスタンス化が発生してから Nothing 代入になるが、プレースホルダーに対して `triggerTerminate` を呼んでも `__vbaClass__` がないため何もしなかった
  - 修正: `Set x = Nothing` の代入処理（`evaluateSetStatement`）で、`oldVal` が auto-instance プレースホルダーのときは `instantiateClass()` で先にインスタンス化してから `triggerTerminate` を呼ぶよう修正

- ✅ **Fix: `New Collection` — キーワードトークンのクラス名が `New` 式で拒否される** | `new-keyword-classname.test.ts`
  - 原因: `New` 式のパーサーが `classNameToken.type !== Identifier` をチェックしており、`Collection` のように VBA 組み込みキーワードとして字句解析されるクラス名を拒否していた
  - 症状: `New Collection` を含む関数全体がエラーリカバリで消滅し、後続のプロシージャが誤ってトップレベルにリークする
  - 修正: `isNameToken()` ヘルパーを追加し、`Identifier` またはキーワード範囲のトークンを許可。`New Scripting.Dictionary` のようなドット付きクラス名も正しく捕捉するよう拡張

- ✅ **Fix: `As VBA.Collection` — モジュール修飾型の戻り値型が半分しかパースされない** | `module-qualified-type.test.ts`
  - 原因: 関数・Subの戻り値型 `As <Type>` のパーサーがトークンを1つだけ読んでいたため、`As VBA.Collection` は `VBA` だけを returnType とし、`.Collection` が本体ストリームに残った
  - 症状: 残った `.Collection`（キーワード）が本体パースでエラーとなり、関数宣言全体が消滅して `[]` になる
  - 修正: `As` 後にドットが続く場合は `Module.Type` 形式で読み切るよう修正。変数宣言の `Dim x As Module.Type` も同様に対応

- ✅ **Fix: `.Collection` — implicit-With のプロパティ名がキーワードのとき失敗する** | `keyword-property-access.test.ts`
  - 原因: `With obj: .Property = x` のパーサーが `propToken.type !== Identifier` をチェックしており、`.Collection` のようにプロパティ名がキーワードトークンのとき例外を投げていた
  - 症状: `With` ブロック内でキーワード名プロパティを使うと、含む Sub/Function 全体がエラーリカバリで消滅する
  - 修正: `isNameToken()` を使い、`Identifier` またはキーワード範囲のトークンを許可

- ✅ **偽陽性: `MySub()` / `MySub(arg)` を文として書いてもエラーにならない** | `tests/spec/vba_compile_error.test.ts`
  - 実 VBA: "ステートメントの末尾が正しくありません" コンパイルエラー
  - 修正: Pass1 で `CallExpression` を `CallStatement` に変換するブランチで、カッコ付き Sub 呼び出しをエラーとするよう修正

- ✅ **偽陽性: `v = MySub` / `v = MySub()` を式文脈で書いてもエラーにならない** | `tests/spec/vba_compile_error.test.ts`
  - 実 VBA: "Function または変数が必要です" コンパイルエラー
  - 修正: Pass2 (`resolveIdentifiers`) で代入右辺・式文脈で Sub を参照した場合に compile error を発生させるよう修正

### Evaluator バグ修正

- ✅ **Fix: モジュールレベル配列の境界式で `Const` 定数を参照すると境界が `upper=-1` になる** | `dim-const.test.ts`
  - 原因: Pass 1（`evaluateModule`）は `ConstDeclaration` のみスキップしていたが、配列境界付き `VariableDeclaration` は即時評価していた。境界式で `env.get("W")` を呼ぶと `W` が未定義のため `vbaEmpty`（null → 0 に数値変換）が返り、`upper = W - 1 = -1` の空配列が生成された
  - 症状: `Const W = 3: Dim a(0 To W - 1) As Integer` の後で `a(2) = 5` が `Subscript out of range (Error 9)` になる
  - 修正: Pass 1 のアーキテクチャを「シンボルテーブル構築のみ」に統一。配列境界付き Dim を `pendingArrayDecls` に退避、モジュールレベル実行文を `pendingTopLevel` に退避し、`resolveIdentifiers`（Pass 2）で Const 確定後にそれぞれ評価・実行する。`resolveIdentifiersDone` フラグで以降の `evaluateModule`（`evalExpression` 等）は即時実行に切り替える

- ✅ **Fix: `With` ブロック内で JS prototype 上の getter/メソッドが解決できない** | `with-statement.test.ts`
  - 原因: `evaluateImplicitWithObjectExpression` が `Object.keys()`（own enumerable プロパティのみ）と `hasOwnProperty()` で検索していたため、TypeScript class の `get accessor` やメソッドといった prototype チェーン上のメンバーを `.Prop` 構文で発見できなかった
  - 症状: `With obj: .Replacement.Text = "bar"` のように、`.Replacement` が prototype 上の `get accessor` として定義されている JS オブジェクトに対して "Object required (424)" が発生する。`obj.Replacement.Text = "bar"` は `evaluateMemberExpression` 経由で prototype を辿るため正常に動く（非対称）
  - 修正: `evaluateImplicitWithObjectExpression` を `resolveObjectMemberKey`（prototype チェーンを辿る）に統一。`evaluateMemberExpression` と同じロジックを使うことで非対称が解消された

- ✅ **Fix: `Not` 演算子・`vbaToBoolean` が JS ネイティブ `boolean` を変換できない** | `excel-stub.test.ts`, `access-stub.test.ts`
  - 原因: `Not` 演算子は `vbaTrue`/`vbaFalse`（VbaBoolean）と数値のみ処理し、JS ネイティブ `true`/`false` に対して `~true = -2`（truthy）を返していた。`vbaToBoolean` も JS boolean を Type mismatch エラーにしていた
  - 症状: JS モックオブジェクトの boolean ゲッター（例: `MockRecordset.EOF = true`）を VBA で `Do While Not rs.EOF` に使うと無限ループになる。`Dim x As Boolean: x = Application.ScreenUpdating` で Type mismatch が発生する
  - 修正: `Not` 演算子に `typeof argument === 'boolean'` ケースを追加。`vbaToBoolean` も同様に JS boolean を `vbaTrue`/`vbaFalse` に変換するよう対応

- ✅ **Fix: `Dim a(n)` 配列サイズに変数を指定してもエラーにならないバグ** | `dim-const.test.ts`
  - VBA 仕様: `Dim` 文の配列境界式は定数式でなければならない。変数参照は "Compile error: Constant expression required" になる
  - 修正: vb* 組み込み定数を `env.setConstant()` で登録し `isConstant()` で識別可能に。`validateConstantExpr()` を追加し Dim 境界を検証
  - モジュールレベル: `resolveIdentifiers()`（Pass 2 resolve）で検出。プロシージャ内: `checkConstantArrayBoundsInProc()`（`precheckProc`）で AST 走査により検出
  - `ReDim` は変数指定可能・実行時評価のまま変更なし

- ✅ **Refactor: `assertCompileErrorPreproc` を `precheckProc` 専用に構造化** | `vba_compile_error.test.ts`
  - 変更前: `throwAction` が `callProcedure`（precheck + 実行の両方を含む）だったため、実行時エラーでもパターンが合えば通過していた
  - 変更後: `checkProcedure()` を public API として公開し、`throwAction` を `checkProcedure` に変更。`precheckProc` のみを実行し本体は実行しない
  - 併せて以下の AST 静的チェックを `precheckProc` に追加（これらが実行時エラーになっていたため）:
    - `checkDuplicateDimInProc`: 同一スコープ内の重複 Dim 宣言
    - `checkGoToLabelsInProc`: GoTo/GoSub の未定義ラベル（`If False Then GoTo` のような未到達コードも検出）
    - `checkCallArgCountsInProc`: ユーザー定義プロシージャへの引数数不一致

- ✅ **Fix: プロシージャの後に書いた Dim/Const/Type/Enum がコンパイルエラーにならない** | `vba_compile_error.test.ts`（`module_level_dim_after_procedure` 等）, `module-level-toplevel-after-procedure.test.ts`
  - VBA 仕様: モジュールレベル宣言（Dim/Const/Type/Enum 等）はすべてのプロシージャより前に書く必要があり、プロシージャ（End Sub/End Function/End Property）の後に書くと「End Sub、End Function または End Property 以降には、コメントのみが記述できます」というコンパイルエラーになる
  - 修正: `resolveIdentifiers`（Pass 2）に、モジュール内で最初の `ProcedureDeclaration` より後に他の文（宣言文・実行文どちらも）が現れたらエラーにするチェックを追加
  - vba-runner 拡張: `evalVBASingle`/`evalVBAModules` 用のテストスクリプトでは、Dim をプロシージャの後に書いたり、For/Next 等の実行文をモジュールレベルに書く書き方が広く使われている（REPL・簡易テスト用）。これを壊さないよう `Evaluator` に `allowTopLevelStatements`（デフォルト `true`）オプションを追加し、true の間は上記チェックを無効化。`false` を指定すると標準 VBA 相当の挙動になる
  - `CompileError.bas` の自動テスト生成器（`compile-error-generator.ts`）に、ケースごとに `Evaluator` のオプションを上書きできる `EVAL_OPTIONS:` メタフィールドを追加。また `RUNNER: TBD` 調査用の `captureCompileErrorMessage()` を追加（既存の `assertCompileErrorPrerun` 等は pattern `/.+/i` に対して常にマッチしてしまい実際のエラーメッセージを取得できていなかったため）

### Parser の拡張機能

- ✅ **Parser に `parseAsClass` パラメーターを追加**: `.cls` ファイルやプログラム的にクラスとしてパースすべきコードの指定 | `parse-as-class.test.ts`
  - `new Parser(tokens, { parseAsClass: 'ClassName' })` でクラスボディとしてパース
  - `parseClassDeclaration` のボディ解析を `parseClassBody(name, untilEndClass)` に抽出し再利用
  - `test-libs/test-runner.ts` の文字列ラップハックを削除し `parseAsClass` オプションに統一

- ✅ **Fix: `Private/Public Type` 宣言が無視される** | `private-type-declaration.test.ts`
  - 原因: パーサーがスコープ修飾子（`Private`/`Public`）の後に `Type` キーワードが来るケースを未処理。`parseDimStatement` にフォールスルーし、`Type` がキーワードトークンのため変数名チェックに失敗してエラーリカバリされていた
  - 症状: `Private Type json_Options ... End Type` が AST に含まれず、`json_Options` 型の変数が常に `0` になる
  - 修正: `Public/Private/Friend` ブロックに `KeywordType` / `KeywordEnum` のケースを追加。`parseTypeDeclaration` の配列メンバー（`name(0 To N) As Type`）も括弧をスキップして対応

- ✅ **Fix: `VBA.vbNull` など `VBA.` モジュール修飾定数アクセスが失敗する** | `vba-module-qualifier.test.ts`
  - 原因: `evaluateMemberExpression` が `VBA.SomeThing` を評価する際、`VBA` 変数が未定義のため `0` になりプロパティアクセスでエラー 424 になっていた
  - 症状: `Case VBA.vbNull` / `Case VBA.vbString` などが実行時エラー 424 で失敗
  - 修正: `possibleModule.toLowerCase() === 'vba'` の場合、プロパティ名を直接 `env.get(propName)` で lookup するフォールバックを追加（`vbNull`, `vbString` 等は既に env に登録済み）

- ✅ **Fix: 再帰関数で `Dim` が外側のスコープを上書きする** | `recursive-dim-scope.test.ts`
  - 原因: `evaluateVariableDeclaration` が `this.env.set(varName, initialValue)` を使っており、`set()` はスコープチェーンを遡って既存の変数を上書きする。再帰呼び出し時に内側の `Dim` が外側の同名変数を初期値でリセットしてしまう
  - 症状: `ConvertToJson` が再帰呼び出しされると、内側の `Dim json_Converted As String` が外側の `json_Converted` を `""` にリセットしてバッファー内容が消失。`{"name":"Alice",...}` が `"":true}` になる
  - 修正: プロシージャ内の変数宣言は `setLocally()` を使用してカレントフレームにのみ変数を作成

- ✅ **Fix: `obj.Method(key) = value` / `Set obj.Item(key) = obj2` / `outer(k1)(k2) = val` 形式の代入が失敗する** | `member-call-assignment.test.ts`, `object-member-assignment.test.ts`
  - 原因: `evaluateAssignmentToVariable` が `CallExpression` の LHS を処理する際、`callee` が `Identifier` の場合のみ対応しており、`callee` が `MemberExpression`（`obj.Item(key)` 形式）の場合にエラー 5 を投げていた
  - 症状: `dict.Item(key) = val` のように `obj.Method(args) = val` 形式で代入すると実行時エラー 5 "Invalid procedure call or argument" が発生
  - 修正: `evaluateAssignmentToVariable` に `callee=MemberExpression`（`obj.Item(k) = v`）と `callee=CallExpression`（`outer(k1)(k2) = v`）のケースを追加
  - `evaluateSetStatement` にも同様の `CallExpression` LHS ケースを追加（`Set obj.Item(k) = obj2`）

- ✅ **Fix: 関数内で `F = F & "X"` が再帰呼び出しになり、戻り値変数の初期値が `0` になる** | `function-return-var.test.ts`
  - 原因1: `Identifier` 評価時に「0引数なら auto-call」するロジックが、現在実行中の関数名（= 戻り値変数）にも適用されていた
  - 原因2: `callProcedure`（公開 API）が `evaluateCallExpression` と異なり、関数の戻り値変数を `vbaEmpty` で初期化していなかった。`env.get` が implicit initialization で `0` を返していた
  - 症状: `F = F & "X"` がスタックオーバーフロー、または戻り値変数が `""` でなく `0` になる
  - 修正: Identifier 評価で `currentProcedureName` と一致する場合は auto-call しない。`callProcedure` に `localEnv.setLocally(proc.name.name, vbaEmpty)` を追加

- ✅ **Fix: `On Error Resume Next` が `Select Case` / `If` / `For` / `While` / `With` ブロック内部で機能しない** | `on-error-in-blocks.test.ts`
  - 原因: `evaluateSelectCaseStatement`・`evaluateIfStatement`・`evaluateForStatement`・`evaluateForEachStatement`・`evaluateDoWhileStatement`・`evaluateWhileStatement`・`evaluateWithStatement` が単純な `for...of` ループでブロック内ステートメントを実行しており、`On Error Resume Next` のエラーキャッチが機能しなかった。エラーはブロック全体を貫通して `executeStatements`（プロシージャ本体レベル）に伝播し、Select Case 全体・If 全体が「1ステートメント」としてスキップされていた
  - 症状: `Select Case VarType(arr)` 内で `LBound(arr, 2)` が VBA エラー 9 を投げると、`On Error Resume Next` で次行に進まず Select Case 全体がスキップされ、戻り値が null になる（`ConvertToJson(array)` が `""` を返す）
  - 修正: 各ブロック評価関数でブロック本体の実行を `executeStatements(body, 0, false)` に変更。`isTopLevel=false` の場合は VBA エラーのみ `On Error Resume Next` で処理し、`GoTo` / `GoSub` / `Return` / `Resume` はプロシージャ本体の `executeStatements` に委譲するよう再スロー

- ✅ **Fix: `VarType(Nothing)` が 9（vbObject）ではなく 12 を返す** | `vartype.test.ts`
  - 原因: `vbaNothing` は `Symbol` であり、`vartype` 組み込み関数の各型チェック（`instanceof`・`typeof object`）をすべてすり抜けて `return 12` にフォールスルーしていた
  - 症状: `JsonConverter.bas` の `json_IsUndefined` が `VarType = vbObject (9)` で Nothing を検出する設計のため、辞書の Nothing 値がスキップされず JSON に出力されてしまう
  - 修正: `vartype` の先頭近くに `if (val === vbaNothing) return 9; // vbObject` を追加

- ✅ **Fix: `Set c = New Collection` / 関数から返した Collection への `For Each` が "Type mismatch" になる** | `for-each.test.ts`
  - 原因: `Dim col As New Collection` は `VbaCollection` クラスインスタンス（`items` getter あり）を生成するが、`Set c = New Collection` は `instantiateClass('Collection')` → `registerExternalObject('Collection', ...)` のファクトリ経由でプレーンオブジェクト（`Symbol.iterator` あり・`items` なし）を生成する。`evaluateForEachStatement` が `items` プロパティの存在のみをチェックしており、ファクトリオブジェクトを処理できなかった
  - 症状: `ParseJson("[1,2,3]")` が返す Collection への `For Each` でエラー 13 が発生
  - 修正: `evaluateForEachStatement` に `__isVbaCollection__` チェックを追加し、`Symbol.iterator` があれば `Array.from` で列挙するパスを設ける

- ✅ **Fix: `Val()` が科学表記（`1.5e2` など）を正しく変換しない** | `conversion-int-val.test.ts`
  - 原因: 正規表現が `/^[+-]?\d*(\.\d*)?/` で `eE` を含まず、`Val("1.5e2")` が `1.5` を返していた
  - 症状: `json_ParseNumber` 内で `VBA.Val("1.5e2")` を呼ぶと 1.5 になり、科学表記の数値が正しくパースされない
  - 修正: 正規表現に `([eE][+-]?\d+)?` を追加

- ✅ **Fix: `vbBack`・`vbFormFeed` 定数が未定義** | `builtin-strings.test.ts`
  - 原因: 定数登録が漏れており、`json_ParseString` の `\b`・`\f` エスケープで `vbBack`/`vbFormFeed` を参照するとエラーになる
  - 修正: `vbBack = "\b"`, `vbFormFeed = "\f"` を定数として登録

- ✅ **Fix: `Hex$`・`Oct$` が未登録（`hex`/`oct` のみあり）** | `number_literals.test.ts`, `builtin-strings.test.ts`
  - 原因: `$` なし版のみ登録しており `Hex$`・`Oct$`・`ChrW$` 呼び出しで Error 438 になっていた
  - 症状: `json_Encode` 内の `"\u" & Right$("0000" & Hex$(AscW), 4)` が失敗し、非ASCII文字の `\uXXXX` エンコードが動かない
  - 修正: `hex$`・`oct$`・`chrw$` を各々のエイリアスとして登録

- ✅ **Fix: `Format$` の `HH`（大文字）が時刻として認識されず、`mm` が文脈判定なしに常に月になっていた** | `format.test.ts`
  - 原因1: `formatDate` の正規表現が小文字 `hh` のみ対応で `HH` をリテラルとして出力していた
  - 原因2: `mm` が月か分かの文脈判定がなく、`"HH:mm:ss"` の `mm` が常に月になっていた
  - 修正: トークン単位の処理に変更し、直前が `h`/`hh` のとき `mm`/`m` を分として解釈する `prevTokenWasHour` フラグを実装

- ✅ **Fix: `Dim ws As Worksheet` など未知の外部オブジェクト型の初期値が `Nothing` でなく `Empty` になる** | `on-error-nothing.test.ts`
  - 原因: 変数宣言の初期化ロジックが `classDefinitions`/`externalObjectFactories` に未登録の型を "Variant や未知の型" と同扱いして `vbaEmpty` を返していた
  - 症状: `Dim ws As Worksheet` で `ws Is Nothing` が False になり、`On Error Resume Next` で `Set` が失敗した後もワークシートアクセスで Error 91 が発生する
  - 修正: `Variant`・`Date` 以外の未知型は外部COMオブジェクト型とみなし `vbaNothing` で初期化

- ✅ **Fix: `Format()` 数値フォーマットの零埋め（`"000"`）が動作せず、また `#` プレースホルダーが `0` と同じ動作になっていた** | `format.test.ts`
  - 原因1: 整数部フォーマットの `0` 個数から最小桁幅を算出して `padStart` する処理が未実装だった
  - 原因2: `#`（有効桁のみ表示）と `0`（常に表示）を区別するロジックがなく、`Format(0, "#")` が `""` でなく `"0"` を返し、`Format(0.5, "#.#")` が `".5"` でなく `"0.5"` を返していた
  - 修正: `minIntegers` を整数部の `0` 個数から算出して `padStart` を追加し、`#` のみのフォーマットで整数部が `"0"` の場合は空文字に抑制するよう修正

- ✅ **Fix: `evaluateDateLiteral` が UTC ベースの Date を生成していたため `formatDate` と時差分ずれていた** | `format.test.ts`
  - 原因: `evaluateDateLiteral` が `Date.UTC(...)` で JS Date を生成していたが、`formatDate` はローカル時刻の `getHours()` 等を使うため時差分ずれていた
  - 修正: `evaluateDateLiteral` を `new Date(y, m-1, d, h, min, s)`（ローカル時刻）に変更し、`Now()`/`DateSerial()` と一貫した挙動にした。`formatDate` はローカル時刻の `get*` メソッドのまま
  - 影響: `ConvertToJson(Date)` の ISO 8601 出力が正しくなった

- ✅ **Fix: `On Error GoTo` ハンドラー内の `If`/`For`/`While` ブロックで `Err.Raise` すると無限ループになる** | `error-handler-reentry.test.ts`
  - 原因: ネストブロックの `executeStatements(isTopLevel=false)` が re-throw 前に `isInErrorHandler = false` にリセットしていたため、外側の `executeStatements` がフラグを見て「ハンドラー外」と誤判断し、同じ GoTo ラベルに再ジャンプしてループ
  - 修正: "bubble up" パスの `this.isInErrorHandler = false` を削除。ネストブロックからの re-throw では `isInErrorHandler` を保持し、外側が正しくバブルアップする
  - 影響: `LibBook.bas` の `NewWorkbook` 関数（Cleanup ハンドラー内で `Err.Raise` するパターン）がハングしなくなった

- ✅ **Fix: `Append`・`Output`・`Binary`・`Random` がファイルモードキーワードとして予約語化され変数名/式に使えない** | `backslash-escape-min.test.ts`
  - 原因: Lexer が `append` / `output` / `binary` / `random` を無条件に `KeywordAppend` 等のトークンに変換していたため、`Dim append As String` の Dim 宣言後の代入（`append = "\\\\"` など）がパースエラーになっていた
  - 症状: `json_BufferAppend` 内の `json_Append` ではなく、ユーザーコードで `append` などを変数名に使うとパース失敗。スタックトレースは "Unexpected token in expression 'append'"
  - 修正: `parseStatementInner` の識別子ステートメント開始条件と `parsePrimary` の文脈的識別子リストに `KeywordOutput`/`KeywordAppend`/`KeywordRandom`/`KeywordBinary` を追加（`Open...For Append` の文法は既に `parseOpenStatement` で正しく処理される）

- ✅ **Fix: `Sub (expr)*x` 形式でスペース前 `(` が関数呼び出しとして貪欲解析され ParseError になる** | `sub-call-paren-expr.test.ts`
  - 原因: `parsePrimary()` のポストフィックスループが `Debug.Print (1+2)` を `CallExpression(Debug.Print, [1+2])` として貪欲消費し、後続の `*3` が新引数の先頭トークンとして解析されようとしたが `*` は式の先頭に置けないため ParseError
  - 症状: `Debug.Print (1+2)*3` が "Unexpected token in expression '*'" でパースエラー。VBA の正しい動作は `9` を表示
  - 修正: `isBinaryOnlyOperator()` で単項不可の二項演算子 (`*`,`/`,`^`,`\`,`+`,`-`,`&`,`Mod`, 比較演算子) を判定。文ステートメントで `parsePrimary()` が `CallExpression` を返し次トークンが該当演算子の場合、位置をリセットして `parsePrimary(stopBeforeSpacedLParen=true)` で再解析。スペース前の `(` をポストフィックス消費せず、`parseCallArgument()` が `(1+2)*3` を一つの引数式として正しく解析する

- ✅ **Parser の EOS 検証を体系化する**（構文精度の根本課題）
  - 背景: MS-VBAL §5.4.1 `statement-block = *(block-statement EOS)` / `EOS = *(EOL / ":")` によれば、各文の後には必ず EOS（改行またはコロン）が必要。
  - 実装: `parseStatement(checkEOS = true)` にデフォルト引数を追加。stmt が非 null かつ `!isAtTerminator()` のとき ParseError を投げる。インライン If の2ループのみ `parseStatement(false)` を渡す（`Else` トークンが EOS の役割を担うため）。
  - 効果: `Return Foo`・`Exit Sub Bar`・`Exit For Extra` 等の余分トークンが全ブロックループで自動的に検出される。`End Foo`・`Stop Foo` の個別チェック（既存）は冗長になるが無害。
  - テスト: `eos-verification.test.ts` | 2026-06-06

- ⚠️ **Lexer のキーワード分類を仕様書 §3.3.5.2 のカテゴリに整理する**（可読性・保守性）
  - 現状: 全キーワードが `Keyword*` トークンのフラットな羅列で、仕様上の種別（`statement-keyword` / `marker-keyword` / `operator-identifier` / contextual）が区別されていない
  - contextual キーワードの誤予約語化（`Append`・`Output` 等が変数名に使えない問題）は Parser の `CONTEXTUAL_KW` Set 化により軽減済み。新規追加も Set に1行追加するだけ
  - 2025-05-30: `Class / Collection / Error / Property` の未登録を修正、`isIdentifier()` ヘルパー導入、`parseVarDeclaration` / `parseConstDeclaration` / `parseProcedureDeclaration` のチェックを統一。詳細は `docs/internals/IDENTIFIER_CLASSIFICATION_TODO.md` を参照
  - 2026-06-06: `parseEnumDeclaration` / `parseTypeDeclaration` / `parseForStatement` / `parseForEachStatementBody` の変数名・型名チェックを `isIdentifier()` に統一。`Next <var>` のチェックも同様に修正 | `contextual-kw-for-enum-type.test.ts`
  - 残課題: キーワード補完実装時に contextual keyword が補完候補に混入する恐れがある
  - 対策案: `TokenType` に `ContextualKeyword*` カテゴリを追加して Lexer 判定を分離

- ✅ **予約語 + 型接尾辞（`dim$`・`for$` 等）が変数名として通ってしまう**（バグ）
  - 修正: Lexer でキーワード照合前に型接尾辞を除去（`lowerBase`）することで `dim$` → `KeywordDim` と正しく識別
  - ただし contextual keyword + `$`（`append$` 等）は IDENTIFIER なので合法 | `lexer-column.test.ts`

- ✅ **Fix: For/ForEach/DoWhile/While ループ本体内のラベルへの GoTo がエラーになる** | `goto-in-loop.test.ts`
  - 原因: `executeStatements(body, 0, isTopLevel=false)` は GoTo 例外を無条件で re-throw する設計で、ループ本体（`stmt.body`）内にあるラベルを自分では探さなかった。`evaluateForStatement` 等が GoTo を受け取っても `Exit For` 以外はすべて上位に投げていたため、トップレベルの `executeStatements` が手続き本体からラベルを探し、ループ内にしか存在しないラベルを見つけられずエラーになっていた
  - 修正: `evaluateForStatement` / `evaluateForEachStatement` / `evaluateDoWhileStatement` / `evaluateWhileStatement` の catch ブロックで GoTo を受け取ったとき、`stmt.body` 内にラベルが存在すれば `executeStatements(stmt.body, labelIndex + 1, false)` でラベル直後から残りを実行するよう変更。`findLabelInBody` ヘルパーを追加
  - 影響: MS-VBAL §5.2.1.2 — ラベルのスコープは手続き全体（ループの内外を問わない）という仕様に準拠

- ✅ **Fix: Option Explicit チェックを呼び出し時の env 状態で判定する** | `cross-module-const.test.ts`
  - 原因: `optionExplicitViolations` に記録された未宣言名を呼び出し前に固定判定していたため、`runner.set()` や別モジュールで後から定義した定数が解決済みと扱われなかった
  - 修正: `callProcedure` 内のチェックを `env.hasVariable(name)` で再評価するよう変更。呼び出し時点の env に名前が存在すれば通過する（`optionExplicitViolations` は `Map<procName, Set<undeclaredNames>>` に変更）
  - 詳細: `docs/internals/NAME_RESOLUTION.md` 参照

- ✅ **Fix: モジュールレベル定数のクロスモジュール参照がロード順依存になっていた** | `cross-module-const.test.ts`
  - 原因: `evaluate()` 実行時に定数の右辺を即評価するため、参照先モジュールがまだロードされていない場合に `env.get()` の暗黙初期化（0）が採用されていた
  - 修正: `VBARunner` に二段階ロードを導入。Pass 1（`evaluateModule()`）で全モジュールをロード後、Pass 2（`resolveIdentifiers()`）で全モジュールレベル定数を再評価。依存グラフをトポロジカルソートして正しい順序で確定させる
  - 仕様: VBA では全モジュールがコンパイル時に一括解決されるため、ロード順によらずクロスモジュール定数参照が動作する（MS-VBAL §5.6.10 tier 4）

- ✅ **Fix: `Private Const` の他モジュールからのアクセスをエラーにする** | `cross-module-const.test.ts`
  - 仕様: VBA では `Private Const` は同一モジュール内からのみ参照可。他モジュールから参照すると「コンパイルエラー 定数式が必要です」
  - 実装: `resolveIdentifiers()` の依存グラフ構築時に、参照先 Const の `scope` が `'private'` かつ別モジュールであればエラーを throw

- ✅ **Fix: モジュールレベル定数の循環参照を検出してエラーにする** | `cross-module-const.test.ts`
  - 原因: 二段階評価の Pass 2 を単純に一度だけ実行した場合、相互参照する定数は不定値になりエラーにもならなかった
  - 修正: `topologicalSortConsts()` で DFS による閉路検出を実装。循環が見つかれば `"Circular reference in constant declarations: A → B → A"` 形式のエラーを throw する
  - 仕様: VBA コンパイラは定数の循環参照をコンパイルエラーとして検出する

- ✅ **Fix: `For i ... Next j` のようにループ変数と Next 後の変数が不一致の場合にコンパイルエラーを生成しない** | `for-next-variable-mismatch.test.ts`
  - 原因: パーサーが `nextIdentifier` を AST に記録するだけで、ループ変数（`identifier` / `variable`）との一致検証を行っていなかった
  - 修正: `parseForStatementBody` と `parseForEachStatementBody` の `Next` 解析直後に大文字小文字を無視した名前比較を追加し、不一致なら `throwError` でコンパイルエラーを throw する。`Next`（変数名なし）は引き続き正常
  - 仕様: VBA コンパイラは「Next で指定された変数の参照が不正です」としてコンパイルエラーを生成する

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

- ❌ **`precheckProc` の AST チェックを 1 パス走査に統合**
  - 現状: `checkSubAsValueInProc`・`checkUndefinedCallsInProc`・`checkConstantArrayBoundsInProc`・`checkDuplicateDimInProc`・`checkGoToLabelsInProc`・`checkCallArgCountsInProc` の 6 メソッドがそれぞれ独立してプロシージャ本体を走査している
  - 改善: 単一の `walkProcBody(visitors)` パスにまとめ、各チェックをビジター関数として渡す設計に統合する

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

- ❌ **識別子の大文字小文字混同の検出**: 同一スコープ内で大文字小文字だけが異なる識別子（変数名・クラス名・プロシージャ名）の宣言を検出してエラーにする
  - 例: `Dim assert As New Assert` — 変数 `assert` とクラス `Assert` は VBA では同一識別子
  - VBA はケースインセンシティブなため、コンパイル時に "識別子が重複しています" とすべきケース
  - 実装案: Lexer/Parser で識別子を正規化（小文字化）した後、同一スコープ内に同名の宣言があれば警告またはエラー

---

## リテラル型・演算子型の未実装項目

詳細は [`docs/internals/LITERALS_AND_OPERATORS.md`](../docs/internals/LITERALS_AND_OPERATORS.md) を参照。

| 状態 | 項目 | 概要 |
|------|------|------|
| ✅ | **数値リテラルのサフィックス型情報保持** | `NumberLiteral` AST の `typeSuffix` フィールドで `TypeName(100&)` → `Long`、`TypeName(1.5!)` → `Single` など正しく返る | `typename.test.ts` |
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

## サンプルコード検証で発見したエンジンバグ（修正済み / 既知制約）

`sample/workspace/` の Shop サンプル（Product.cls / Inventory.cls / ShopForm.cls / ShopMain.bas）を実行した際に発見。

### ✅ B-1: `parseAsClass` 時にクラス内 `Private Const` が実行時に 0 として解決される（修正済み）

- **修正**: `parser.ts` の `parseClassBody` に `KeywordConst` ケースを追加。`evaluator.ts` の `createInstanceFromDef` でクラスレベルの Const 宣言を `instanceEnv` に評価・登録。
- | `parse-as-class.test.ts` (Test 9: parseAsClass: B-1 — Private Const クラス内参照)

### ✅ B-2: クラス内プライベートメソッドを同クラスメソッドから呼べない（修正済み）

- **修正**: `evaluateCallExpression` で `getProcedure` より先に `Me.__classDef__.procedures` を検索する。VBA の正しい名前解決順序（クラス自身のスコープ → グローバル）を実装。
- | `parse-as-class.test.ts` (Test 10: B-2 — クラス内プライベートメソッド呼び出し, Test 11: B-2 — クラス自身のスコープがグローバルより優先)

### ✅ B-3: `Class_Terminate` スコープ終了時の自動発火（`New` 追跡方式）

- **実装**: `Set x = New ClassName` で作成したインスタンスを `_currentNewOwned` に登録し、スコープ終了時（Sub/Function/クラスメソッドの finally）に `Class_Terminate` を自動発火する。
- **既知の制限 (⚠️ 真の参照カウントではない)**:
  - `Set x = New ClassName` で直接作成したインスタンスのみ追跡。ファクトリ関数経由（`Set x = CreateMyClass()`）で取得したオブジェクトは追跡されず、明示的な `Set x = Nothing` が必要。
  - Dictionary/Collection に格納後もローカル変数に残っているオブジェクトは、スコープ終了時に Terminate が呼ばれる（Dictionary が参照を保持していても）。真の COM 参照カウント方式では Terminate は起きない。
  - 外部スコープから借用したオブジェクト（`Set ws = Workbook.Sheets("名前")` 等）は `New` 式ではないため追跡されず、早期 Terminate しない（安全）。
- | `circular-reference-terminate.test.ts` (Test 11: B-3 — Dictionary 格納済みオブジェクトの早期 Terminate が発生しない, Test 12〜14: スコープ終了時の自動 Terminate)

### ✅ B-4: UDT 固定長配列要素へのフィールド代入が Error 91（修正済み）

- **修正**: `evaluator.ts` に `fillArrayWithUdtInstances` を追加。固定サイズ UDT 配列の各要素を `instantiateType` で個別に初期化する。
- | `udt-type.test.ts` (Test 8: B-4 — 各要素が独立したインスタンス, Test 9: B-4 — 要素間で参照が共有されていない)

### ✅ B-5: クラス内で `Me.Property = value` が Error 91（修正済み）

- **修正**: `parser.ts` の `parseStatement` で識別子ブランチに `TokenType.KeywordMe` を追加。`Me.Prop = val` 形式の代入文が `ImplicitWithObjectExpression`（`.Prop = val`）と誤認識されていたのを修正。
- | `parse-as-class.test.ts` (Test 12: B-5 — Me.Property 代入がクラス内で動作する)

### ✅ B-6: クラス内 `Private WithEvents` フィールド宣言が無視される（修正済み）

- **修正**: `parser.ts` の `parseClassBody` で `KeywordWithEvents` トークンをフィールド宣言開始として認識するよう条件を拡張。`Private WithEvents m_inv As Inventory` のような宣言がフィールドとして登録されず、`Class_Initialize` 内の `Set m_inv = New Inventory` がインスタンス環境に反映されなかった。
- | `parse-as-class.test.ts` (Test 13: B-6 — Private WithEvents フィールドが Class_Initialize で初期化される)

### ✅ B-7: `Event` / `RaiseEvent` でキーワードをイベント名として使えない（修正済み）

- **修正**: `parser.ts` の `parseEventDeclaration` / `parseRaiseEventStatement` で `isNameToken` チェックを追加。`Public Event Open()` / `RaiseEvent Open` のように VBA キーワードをイベント名として使うケース（Excel の Workbook.Open イベント等）が "Expected identifier" エラーになっていた。
- | `raiseevent.test.ts` (既存テストが回帰なしで通過することを確認)

### ✅ B-8: `Function Open/Close()` はモジュールレベルでは仕様上コンパイルエラー（修正済み）

- **経緯**: 過去に `Function Open()` / `Function Close()` をプロシージャ名として受け入れる実装（B-8 fix）を行ったが、Open / Close は `statement-keyword` = `reserved-identifier` （§3.3.5.2）のためモジュールレベルの宣言名には使えない。
- **修正**: `parseProcedureDeclaration` に `STATEMENT_KW_RESERVED` チェックを追加し、これらをコンパイルエラーとして正しく reject するよう変更。クラスモジュールのメンバーとしての `obj.Open` / `obj.Close` はメンバーアクセス（unrestricted-name）なので引き続き有効。
- | `contextual-keyword-as-identifier.test.ts` (Section 6: statement-keyword をプロシージャ名に使うとコンパイルエラー)

### ✅ B-9: クラスの Private/Public フィールドの既定値が型によって Empty のままになる（2026-06-26 修正済み）

- **症状①（UDT）**: `Private mStats As StatBlock`（UDT 型）のようなクラスフィールドが既定値 `Empty` のまま初期化され、`Class_Initialize` 内で `mStats.Strength = 5` のようなメンバー代入をすると `Error 91: Object variable or With block variable not set` になる
- **症状②（Boolean 等）**: `Public IsActive As Boolean` が既定値 `False`(0) ではなく `Empty` になり、`If f.IsActive Then` が意図通り動かない。`Currency`/`Byte`/`LongLong`/`LongPtr` も同様に `0` ではなく `Empty` のままだった
- **症状③（クラス名型フィールド）**: `Public Item As Pub`（クラス型フィールド）が既定値 `Nothing` ではなく `Empty` になり、`Is Nothing` が常に `False` を返す
- **原因**: `Dim` 変数の既定値初期化（`evaluateVariableDeclaration`）は `instantiateType()` での UDT 対応や `boolean`/`currency`/`byte`/`longlong`/`longptr`/クラス型（→Nothing）を含む型集合を持つが、クラスフィールドの初期化（`createInstanceFromDef`）は別の独立したコードパスで、`string`/`integer`/`long`/`double`/`single` の数種類しか default 値を設定していなかった
- **修正**: `createInstanceFromDef` のフィールド既定値判定を `evaluateVariableDeclaration` と同じ型集合に揃え、UDT 型は `instantiateType()` で、クラス名型は `vbaNothing` で初期化するよう変更（`Date`/`Variant` は `Dim` 側でも `Empty` のままが既定の挙動のため対象外）
- | `udt-type.test.ts` (Test 10〜12), `class-module.test.ts`（Boolean/Currency/Byte フィールドの既定値）, `class-object-field-nothing-default.test.ts`

### ✅ B-10: `Set obj.Field = New X` での WithEvents フィールドへの外部代入でイベントハンドラーが配線されない（2026-06-26 修正済み）

- **症状**: `Set s.Source = New Pub`（クラス外部からのメンバーアクセス形式の代入）で WithEvents フィールドに代入しても、イベントハンドラーが一切ワイヤリングされず、イベントが静かに発火しない。`s.Source Is Nothing` は正しく `False`（代入自体は成功している）だが、`RaiseEvent` を含む操作をしても `Source_Ping` ハンドラーが呼ばれない
- **原因**: `evaluateSetStatement` の WithEvents バインディングロジックが `Identifier` LHS ブランチ（`Set Source = New Pub` の形）にしか存在せず、`MemberExpression` LHS ブランチ（`Set s.Source = New Pub` の形）には対応するコードがなかった
- **修正**: バインディングロジックを `bindWithEventsHandlers()` ヘルパーに抽出し、`MemberExpression` LHS ブランチでも `instanceEnv.isWithEvents(propName)` を確認して配線するよう拡張。`classDef`/`instance` にはアサイン先のオブジェクト自身（`obj`）を渡す
- | `raiseevent.test.ts` (Test 5: member access 経由の WithEvents フィールド代入でもハンドラーが配線される)

### ✅ Bug 29-A〜H: 評価 #29 で発見された 8 件のバグ（2026-07-17 修正済み）

- ✅ **Bug 29-A: 型宣言サフィックス付き `Dim n&` / `Dim s$` 等が機能しない**
  - 修正: `parser.ts` の `parseDimStatement` でサフィックス（`%&!#@$^`）を identifier から除去し、型情報を `objectType` に変換。`parsePrimary` でも同じサフィックスを除去して変数参照が一致するよう修正
  - | `type-system.test.ts` (Bug 29-A: 型宣言サフィックス)

- ✅ **Bug 29-B: `^` 演算子が右結合（VBA は左結合）**
  - 修正: `parser.ts:parseExponentiation` の右辺を `parseUnary()`（→ `parseExponentiation` 再入）から単項演算子のみ処理する直接 `parsePrimary()` に変更。`2 ^ 3 ^ 2` = 64 になる
  - | `operators-extra.test.ts` (Bug 29-B: ^ 左結合)

- ✅ **Bug 29-C: 負の底の非整数べき乗が `NaN` を静かに返す（Error 5 にならない）**
  - 修正: `evaluator.ts:case '^'` に `isNaN(result)` チェックを追加し、NaN の場合は Error 5 を throw
  - | `operators-extra.test.ts` (Bug 29-C: (-8)^0.5 → Error 5)

- ✅ **Bug 29-D: `Null & Null` が `""` を返す（Null を返すべき）**
  - 修正: `evaluator.ts:op === '&'` ブランチに「両辺ともに Null なら即 Null を返す」チェックを追加。片側のみ Null の場合は従来通り `""` 扱い
  - | `operators-extra.test.ts` (Bug 29-D: Null & Null → Null)

- ✅ **Bug 29-E: `Array()` が `Option Base 1` を無視する（常に LBound=0 を返す）**
  - 修正: `builtins.ts:array` 登録で `ctx.arrayBase` 分の filler スロットを先頭に追加（`UBound = length - 1` に合わせて `UBound` の計算式も `vbaBase + a.length - 1` → `a.length - 1` へ修正）
  - | `option_base.test.ts` (Bug 29-E: Array() が Option Base 1 を反映する)

- ✅ **Bug 29-F: `eval('a.Increment')` が silent no-op（Sub が実行されない）**
  - 修正: `evaluator.ts:evaluateMemberExpression` の「引数なし Proc 呼び出し」検索条件から `&& p.isFunction` を除去（Sub も呼ぶようにする）
  - | `class-module.test.ts` (Bug 29-F)

- ✅ **Bug 29-G: `eval('Dim g As New Counter : g.Increment : ...')` が Error 450**
  - 修正: `parser.ts` のコール文引数チェックから `TokenType.OperatorColon` を除外（`:` 文区切りを引数区切りと誤認識して `MissingArgument` を生成していた）
  - | `class-module.test.ts` (Bug 29-G)

- ✅ **Bug 29-H: VERSION 行のみで BEGIN/END なしの `.cls` ファイルが本体を全消去する**
  - 修正: `preprocessor.ts:stripVBAFileHeader` で BEGIN 行がある場合のみ BEGIN/END ブロックを除去するよう修正
  - | `preprocessor-cls-header.test.ts` (Bug 29-H)

### ✅ Bug CM〜CT: 監査バッチ（2026-07-17 修正済み）

- ✅ **Bug CM: `Val(42)` が 0 を返す（非文字列引数を String 強制変換してから解析すべき）**
  - 修正: `val` 実装で `typeof s !== 'string'` のとき `vbaToString(s)` で変換してから数値解析
  - | `conversion-int-val.test.ts`

- ✅ **Bug CN: `StrComp("a", "b", Null)` がエラーにならず Binary 比較結果を返す**
  - 修正: Compare=Null のとき `ctx.throwError(VbaErrorCode.INVALID_USE_OF_NULL, ...)` (error 94)
  - | `strcomp.test.ts`

- ✅ **Bug CO: `Switch(False, "A", True)` が undefined を返す（奇数個引数で Error 5 を投げるべき）**
  - 修正: `args.length % 2 !== 0` のとき `ctx.throwError(VbaErrorCode.INVALID_PROCEDURE_CALL, ...)` (error 5)
  - | `choose_switch.test.ts`

- ✅ **Bug CP: `Round(1.5, Null)` が JS TypeError でクラッシュする（Error 13 を投げるべき）**
  - 修正: `digits === vbaNull` のとき `ctx.throwError(VbaErrorCode.TYPE_MISMATCH, ...)` (error 13)
  - | `builtins.test.ts`

- ✅ **Bug CQ: `Join("notanarray", ",")` が String 化・`Join(Array(1, Null, 3), ",")` が JS TypeError**
  - 修正: 非配列引数は error 13、Null 要素も error 13 に統一
  - | `split-join.test.ts`

- ✅ **Bug CR: `DateAdd("zz", 1, ...)` が無効 Interval でエラーなし（Error 5 を投げるべき）**
  - 修正: else 節に `ctx.throwError(VbaErrorCode.INVALID_PROCEDURE_CALL, ...)` を追加
  - | `dateadd-monthend-rollover.test.ts`

- ✅ **Bug CS: `DateDiff("zz", d1, d2)` が無効 Interval で 0 を返す（Error 5 を投げるべき）**
  - 修正: else 節に `ctx.throwError(VbaErrorCode.INVALID_PROCEDURE_CALL, ...)` を追加
  - | `dateadd-monthend-rollover.test.ts`

- ✅ **Bug CT: `DatePart("zz", d)` が無効 Interval で 0 を返す（Error 5 を投げるべき）**
  - 修正: else 節に `ctx.throwError(VbaErrorCode.INVALID_PROCEDURE_CALL, ...)` を追加
  - | `dateadd-monthend-rollover.test.ts`

- ✅ **Bug CU: `DateSerial(10, 1, 1)` が year 1910 を返す（仕様: 0-29 → 2000-2029）**
  - 修正: `new Date()` 呼び出し前に 2桁年を4桁年に変換（JS legacy behavior 回避のため `setFullYear()` も追加）
  - | `datetime.test.ts`

- ✅ **Bug CV: `Weekday("2025-01-01", Null)` が JS TypeError でクラッシュ（VBA Error に変換すべき）**
  - 修正: `firstdayofweek === vbaNull` のとき `ctx.throwError(VbaErrorCode.INVALID_PROCEDURE_CALL, ...)` (error 5)
  - | `datetime.test.ts`

- ✅ **Bug CX: `Hex(-1)` が "FFFFFFFF"（32ビット）を返す（仕様: "FFFF"、16ビット4文字）**
  - 修正: -32767〜-1 の範囲は `(n & 0xFFFF).toString(16).padStart(4)` で16ビット処理
  - | `conversion-int-val.test.ts`

- ✅ **Bug CY: `Oct(-1)` が "37777777777"（32ビット）を返す（仕様: "177777"、16ビット6文字）**
  - 修正: -32767〜-1 の範囲は `(n & 0xFFFF).toString(8).padStart(6)` で16ビット処理
  - | `conversion-int-val.test.ts`

- ✅ **Bug CZ: `StrReverse(Null)` が vbaNull を返す（仕様: エラー発生）**
  - 修正: Null 入力で `ctx.throwError(VbaErrorCode.INVALID_USE_OF_NULL, ...)` (error 94)
  - | `strreverse.test.ts`

- ✅ **Bug DA: `IsNumeric(#1/1/2000#)` が True を返す（仕様: False — §6.1.2.7.1.8 は Date を数値型リストに含めない）**
  - 修正: `builtins.ts` の `isnumeric` から `|| val instanceof VbaDate` を削除
  - | `builtins.test.ts`

- ✅ **Bug DB: `String(3, 257)` が "āāā" を返す（仕様: §6.1.2.11.1.38 — 数値 > 255 は `character Mod 256` を使用 → chr(1)*3）**
  - 修正: `String.fromCharCode(char)` → `String.fromCharCode(Math.trunc(char) % 256)`
  - | `builtin-strings.test.ts`

- ✅ **Bug DC: `String(3, "")` がエラーを発生しない（仕様: §6.1.2.11.1.38 — 空文字 Character は Error 5）**
  - 修正: 空文字列チェック後 `ctx.throwError(VbaErrorCode.INVALID_PROCEDURE_CALL, ...)`
  - | `builtin-strings.test.ts`

- ✅ **Bug DD: `LBound(arr, Null)` / `UBound(arr, Null)` が JS TypeError でクラッシュ（VBA Error に変換すべき）**
  - 修正: `if (dim === vbaNull) ctx.throwError(VbaErrorCode.INVALID_USE_OF_NULL, ...)`
  - | `builtins.test.ts`

- ✅ **Bug DE: `LBound(arr, 0)` / `UBound(arr, 0)` が Error 9 を発生しない（次元は1ベース、0は無効）**
  - 修正: `if (dimIndex > 0)` → `if (dimIndex < 0 || dimIndex > 0)` で負値も捕捉
  - | `builtins.test.ts`

- ✅ **Bug DG: `StrConv("abc", Null)` が JS TypeError でクラッシュ（Conversion は VbStrConv 型、Null → Error 94）**
  - 修正: `if (conv === vbaNull) ctx.throwError(VbaErrorCode.INVALID_USE_OF_NULL, ...)`
  - | `strconv_null.test.ts`
