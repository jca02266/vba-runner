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

## 言語実装の拡張機能

| 状態 | 優先度 | 機能 | 概要 | テスト |
|------|--------|------|------|--------|
| ✅ | P2 | 日本語識別子対応 | MS-VBAL §3.3.5 に従い `isAlpha()` を Unicode プロパティエスケープ (`\p{L}`) で拡張。パーサーの `isWordToken()` も同様に対応。`Dim 氏名 As String` 等の識別子・UDT・関数名で使用可能に | `unicode-identifiers.test.ts` |

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
| ✅ | P1 | TypeName | §6.1.2.3.1.22 | (制限事項: 全ての数値はJSのnumber(Double)として扱われ、Integer/Longの区別は未実装) | `typename.test.ts` |

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
| ✅ | P1 | Filter | §6.1.2.11.1.7 | `filter.test.ts` |
| ✅ | P0 | Format / Format$ | §6.1.2.11.1.8/9 | `format.test.ts` |
| ✅ | P0 | InStr / InStrB | §6.1.2.11.1.14 | `builtin-strings.test.ts` |
| ✅ | P1 | InStrRev | §6.1.2.11.1.15 | `instrrev.test.ts` |
| ✅ | P1 | Join | §6.1.2.11.1.16 | `split-join.test.ts` |
| ✅ | P1 | LCase / LCase$ | §6.1.2.11.1.17/18 | `builtin-strings.test.ts` |
| ✅ | P1 | Left / Left$ | §6.1.2.11.1.19/20 | `builtin-strings.test.ts` |
| ✅ | P0 | Len | §6.1.2.11.1.22 | `builtin-strings.test.ts` |
| ✅ | P1 | LTrim / RTrim | §6.1.2.11.1.23 | `trim.test.ts` |
| ✅ | P1 | Trim / Trim$ | §6.1.2.11.1.23 | `trim.test.ts` |
| ✅ | P0 | Mid / Mid$ | §6.1.2.11.1.25/26 | `builtin-strings.test.ts` |
| ✅ | P1 | Replace | §6.1.2.11.1.29 | `builtins.test.ts` |
| ✅ | P1 | Right / Right$ | §6.1.2.11.1.30/31 | `builtin-strings.test.ts` |
| ✅ | P1 | Space / Space$ | §6.1.2.11.1.33/34 | `builtin-strings.test.ts` |
| ✅ | P1 | Split | §6.1.2.11.1.35 | `split-join.test.ts` |
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
| ✅ | P0 | TypeName | §6.1.2.7.1.12 | (UDTおよびClassに対応) | `typename.test.ts` |
| ✅ | P1 | VarType | §6.1.2.7.1.13 | `vartype.test.ts` |
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
    - ❌ **1D 配列での単一行/列設定**: VBA では `Range("A1:A4").Value = Array(10,20,30,40)` のように 1D 配列で単一列（行）に書き込めるが、`MockWorksheet.setCellValue` は 2D 配列のみ対応。
    - ❌ **配列サイズ不一致時のエラー検出**: 範囲と配列のサイズが合わない場合、VBA は Error 1004 を発生させるが、MockWorksheet はサイレントにスキップ/空文字填充する。
    - ❌ **`Range().Value =` の書き戻し**: VBA では `ws.Range("A1:B3").Value = array` でセルに書き込めるが、MockWorksheet では `setCellValue` を使わないとセルに反映されない。
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

- ❌ **Range 変数経由のアクセスは未検出**: `rng.Item(3, 5)` / `rng(3, 5)` のように、Range 型の変数を介したアクセスは型追跡（データフロー解析）が必要なため検出できない。`Dim rng As Range` 宣言と `Set rng = ...` 代入を辿って型を推定する仕組みが必要。

## 仮想ファイルシステム (VFS) の実装状況と課題

Webブラウザおよびテスト環境向けの仮想ファイルシステム (`MemoryFileSystem`) の実装状況です。

### 機能マトリックス

| 状態 | 機能カテゴリ | 項目 | 説明 |
| :---: | :--- | :--- | :--- |
| ✅ | **基本 I/O** | `readSync` / `writeSync` | バッファー/文字列の同期読み書き。ハンドル位置の自動更新対応。 |
| ✅ | | `openSync` / `closeSync` | ファイルオープン。`w` フラグ時の truncation (切り詰め) 対応済み。 |
| ✅ | **ファイル管理** | `unlinkSync` | `Kill` ステートメントの基盤。現状は完全一致のみ。 |
| ✅ | | `copyFileSync` | `FileCopy` ステートメントの基盤。 |
| ✅ | **ディレクトリ** | `mkdirSync` | ディレクトリ作成。`recursive: true` 対応。 |
| ✅ | | `readdirSync` | ディレクトリ一覧取得。`Dir` 関数の基盤。 |
| ✅ | | `rmdirSync` | ディレクトリ削除。 |
| ✅ | **メタデータ** | `statSync` | `size`, `mtime`, `birthtime`, `mode` の取得。 |
| ✅ | | `existsSync` | ファイル/ディレクトリの存在確認。 |
| ✅ | **高度な操作** | **ワイルドカード** | `Kill` および `Dir` における `*`, `?` のサポート。 |
| ⚠️ | | **カレントディレクトリ** | `ChDir` / `CurDir` による仮想的な作業ディレクトリの保持。 |
| ❌ | | **永続化** | `localStorage` や `IndexedDB` への保存・復元。 |
| ❌ | | **排他制御** | `Lock` / `Unlock` ステートメントのエミュレーション。 |
| ⚠️ | **互換性** | **バイナリ/テキスト** | `Binary` / `Random` / `Input` / `Output` 各モードの厳密な挙動。 |

### VFS 開発ロードマップ (TODO)

#### 1. コア機能の強化
- ✅ `MemoryFileSystem` にワイルドカードマッチングロジックの実装（`Kill` への統合）。 | `wildcard-kill-dir.test.ts`
- ✅ `MemoryFileSystem` 内部での `cwd` (Current Working Directory) の保持と `ChDir` 対応。 | `chdir-curdir.test.ts`
- ❌ `statSync` における VBA 属性（Read-only, Hidden 等）のシミュレーション。

#### 2. ブラウザ環境最適化
- ❌ `FileSystem` インターフェースの非同期版 (`read`, `write` 等) の検討（ブラウザのメインスレッドをブロックしないため）。
- ❌ `IndexedDB` をバックエンドとした `PersistentFileSystem` の実装。

#### 3. テスト環境の改善
- ❌ `filesystem-extra.test.ts` 等のテストコードを、Node.js `fs` 直接参照から `this.fs` (抽象インターフェース) 参照へリファクタリング。
- ❌ 複数の `Evaluator` インスタンス間で単一の `MemoryFileSystem` を共有するためのテスト用設定の追加。

#### 4. バイナリ操作の極致
- ❌ `Put` / `Get` における固定長文字列、UDT、多次元配列のバイナリレイアウトの完全な VBA 互換性検証。

---

## VBA ランタイム挙動

仕様書本文には書かれているが、個別セクションを持たないため見落とされやすい **暗黙のランタイム挙動** の実装状況をトラッキングするセクションです。チェックなしの項目は **未実装 または 未検証** を意味します。実装する際は「ロードマップ」の手順に従い、VBA ベテランが書きそうな代表コードをテストに含めてください。

### オブジェクトのライフサイクルと参照

- ✅ **Auto-Instantiation (`Dim x As New ClassName`)**: 遅延インスタンス化 + `Set = Nothing` 後の自動再生成 + `Is Nothing` 常時 False
  - ✅ **Auto-Instanceを関数/メソッド引数として渡す**: 未アクセスのAuto-Instanceを引数に渡す際にプレースホルダーのまま渡されるバグを修正 | テスト: `auto-instance-args.test.ts`
    - 修正箇所: JS組み込み関数引数 (L4611)、VBAクラスメソッド引数 (L4698/L4704)、JSオブジェクトメソッド引数 (L4727)、汎用フォールバック (L4752)、デフォルトプロパティ引数 (L4641)
    - 実害バグ: `CallByName body, "Method", VbMethod, autoInstance`（`__vbaClass__` 直接参照で失敗）および `d.Add b, val` でbをキーに使う場合（後でbが解決されると`Exists`がFalse）
- ✅ **`Dim x As ClassName`（New なし）のデフォルト値**: `Nothing` 初期化
- ✅ **Default Property / Default Member**: `Range("A1") = 10` のような暗黙の `.Value` 解決、Collection の `Item` 暗黙呼び出し等
  - ✅ 包括的なテストスイート作成 (`default-property.test.ts`, 15テスト全パス)
  - ✅ 基盤研究: VBA 仕様書からの解析、実装パターン検討
  - ✅ 実装完了: (1) obj = value -> obj.Value = value (Test 14パス)
  - ✅ 実装完了: (2) obj(args) -> obj.Item(args) (Test 12-13パス)
  - ✅ 実装完了: (3) result = obj -> obj.Value (Test 15パス) - 値コンテキストでの暗黙 Value getter
  - ✅ 修正: Test 7 ByRef パラメーター - implicit Value getter の実装で副次的に修正
  - ✅ **非 `__vbaClass__` モックオブジェクトのデフォルトプロパティ**: `x = ws.Range("A1")` で MockRange の `.Value` が自動抽出される
    - opt-in 方式: モックオブジェクトに `__vbaDefault__ = true` と `valueOf()` を実装することで有効化
    - `VbaDate` / `VbaBoolean` / `VbaErrorValue` 等の内部型は `__vbaDefault__` を持たないため誤抽出しない
    - `MockRange` はすでに対応済み (`__vbaDefault__ = true`, `valueOf()` 実装)
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
- ⚠️ **`Decimal` 型の精度**: 28 桁の固定精度サポート (制限事項: VbaDecimalラッパー存在; 内部はJavaScriptのnumber型でIEEE 754精度、28桁固定小数点精度未実装) | `decimal-precision.test.ts`
- ✅ **`Currency` 型の精度**: 64-bit 整数 / 10000 の固定小数 | `currency-precision.test.ts`

### Variant 配列とコレクション

- ✅ **2 次元 Variant 配列の `For Each` 挙動**: 行優先 vs 列優先の反復順 | `2d-array-foreach.test.ts`
- ✅ **Array() で作る Variant 配列の境界**: `LBound = 0`、明示的 `Option Base` の効果 | `array-lbound-defaults.test.ts`
- ✅ **`ReDim Preserve` の多次元配列での制約**: 最終次元のみ変更可 | `redim-preserve-multidim.test.ts`
- ✅ **UDT 配列の `ReDim` 時の要素初期化**: `Dim d() As MyType` + `ReDim d(0 To N)` で各要素が UDT インスタンスではなく `0` になるバグを修正 | `udt-type.test.ts`

### エラーハンドリング

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

### テストランナー (`VBARunner`) の改善

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

- ⚠️ **Lexer のキーワード分類を仕様書 §3.3.5.2 のカテゴリに整理する**（可読性・保守性）
  - 現状: 全キーワードが `Keyword*` トークンのフラットな羅列で、仕様上の種別（`statement-keyword` / `marker-keyword` / `operator-identifier` / contextual）が区別されていない
  - contextual キーワードの誤予約語化（`Append`・`Output` 等が変数名に使えない問題）は Parser の `CONTEXTUAL_KW` Set 化により軽減済み。新規追加も Set に1行追加するだけ
  - 残課題: キーワード補完実装時に contextual keyword が補完候補に混入する恐れがある
  - 対策案: `TokenType` に `ContextualKeyword*` カテゴリを追加して Lexer 判定を分離

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

- ✅ **`Identifier` AST ノードへの `loc` 付与 + スコープ対応シンボルテーブル**: 宣言位置の Identifier に `loc` を付与し、LSP のシンボル参照をスコープ対応に刷新
  - `parser.ts`: `makeIdentifier(token)` ヘルパー追加；Dim・Sub 名・For 変数・Const など全宣言位置に正確な `loc`
  - `symbol-table.ts`: フラット Map → `ScopedSymbolTable`（手続きスコープ別 `localSymbols`）+ `lookupSymbol()`（内側優先解決）
  - 効果: 同名ローカル変数が別 Sub にある場合でも「定義へ移動」「参照検索」「ホバー」「リネーム」が正しいスコープを対象にする
  - `extension.ts` / `code-lens-provider.ts`: `findReferences` コマンドの引数バグ修正（`procName` → `line, character`）

### VBA 仕様制約の検証

- ✅ **モジュール名の長さ検証（31 文字制限）**: MS-VBAL §5.2 で定義されたモジュール名の最大長を実行時に検証
  - 仕様: モジュール名は最大 31 文字（ファイル名またはVB_Name属性）
  - 実装: `Evaluator.setSourceModule()` でモジュール名長をチェック、超過時にエラー発生
  - テスト: `module-name-length.test.ts`, `module-name-length-integration.test.ts`

- ❌ **静的 vs 動的名前解決の区別が不足**: プロシージャ呼び出し時のエラー検出タイミング
  - 優先度: **低** （Option Explicit あり時は問題なし）
  - 実VBA動作の違い：
    - `unknownProc()` → **コンパイルエラー**（Sub/Function が定義されていません）
    - `unknownModule.unknownProc()` → **実行時エラー 424**（オブジェクトが必要です）
  - 根本原因: VBA仕様では修飾/非修飾で名前解決方式が異なる
    - 非修飾: 静的検証（コンパイル時）
    - 修飾: 動的解決（実行時のメンバーアクセス）
  - 当エンジン問題: 両者を同じレベルで処理している（4199-4203行の try-catch が根本原因）
  - 注記: **Option Explicit がない場合の未定義変数の動作**に限定される問題。モダンなVBA開発では Option Explicit が標準であるため、この互換性追求の優先度は低い。Option Explicit ありの場合はコンパイル時に検出される。
  - 修正内容: Parser/Evaluator で修飾/非修飾を区別し、エラーのタイミングを正確に再現

- ❌ **識別子の大文字小文字混同の検出**: 同一スコープ内で大文字小文字だけが異なる識別子（変数名・クラス名・プロシージャ名）の宣言を検出してエラーにする
  - 例: `Dim assert As New Assert` — 変数 `assert` とクラス `Assert` は VBA では同一識別子
  - VBA はケースインセンシティブなため、コンパイル時に "識別子が重複しています" とすべきケース
  - 実装案: Lexer/Parser で識別子を正規化（小文字化）した後、同一スコープ内に同名の宣言があれば警告またはエラー

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
