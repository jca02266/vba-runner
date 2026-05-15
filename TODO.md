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

## 第5章：制御ステートメント (§5.4.2)

| 状態 | 優先度 | 機能 | 仕様書 | テスト |
|------|--------|------|--------|--------|
| ✅ | P0 | Call Statement | §5.4.2.1 | `call-named-params.test.ts` |
| ✅ | P1 | Named Parameters | §5.6.3.1.2 | `call-named-params.test.ts` |
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
| ✅ | P1 | User Defined Type Declarations（`Type`） | §5.2.3.3 | `udt-type.test.ts` |
| ✅ | P1 | Enum Declarations | §5.2.3.4 | `enum.test.ts` |
| ✅ | P0 | External Procedure Declaration（`Declare`） | §5.2.3.5 | (制限事項: JSランタイム上ではスタブとして登録) | `declare.test.ts` |
| ✅ | P0 | Class Module（OOP） | §5.2.4 | `class-module.test.ts` |
| ✅ | P1 | Implements Directive | §5.2.4.2 | `implements.test.ts` |
| ✅ | P2 | Event Declaration | §5.2.4.3 | `raiseevent.test.ts` |
| ✅ | P1 | Option Explicit（パース済み、実行時は無視） | §5.2 | `math-module.test.ts` |
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

本プロジェクトの性質上、テストの安定性および検証精度の向上のために以下の実装を今後検討します。

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

## 仮想ファイルシステム (VFS) の実装状況と課題

Webブラウザおよびテスト環境向けの仮想ファイルシステム (`MemoryFileSystem`) の実装状況です。

### 機能マトリックス

| 状態 | 機能カテゴリ | 項目 | 説明 |
| :---: | :--- | :--- | :--- |
| ✅ | **基本 I/O** | `readSync` / `writeSync` | バッファ/文字列の同期読み書き。ハンドル位置の自動更新対応。 |
| ✅ | | `openSync` / `closeSync` | ファイルオープン。`w` フラグ時の truncation (切り詰め) 対応済み。 |
| ✅ | **ファイル管理** | `unlinkSync` | `Kill` ステートメントの基盤。現状は完全一致のみ。 |
| ✅ | | `copyFileSync` | `FileCopy` ステートメントの基盤。 |
| ✅ | **ディレクトリ** | `mkdirSync` | ディレクトリ作成。`recursive: true` 対応。 |
| ✅ | | `readdirSync` | ディレクトリ一覧取得。`Dir` 関数の基盤。 |
| ✅ | | `rmdirSync` | ディレクトリ削除。 |
| ✅ | **メタデータ** | `statSync` | `size`, `mtime`, `birthtime`, `mode` の取得。 |
| ✅ | | `existsSync` | ファイル/ディレクトリの存在確認。 |
| ⚠️ | **高度な操作** | **ワイルドカード** | `Kill` および `Dir` における `*`, `?` のサポート。 |
| ⚠️ | | **カレントディレクトリ** | `ChDir` / `CurDir` による仮想的な作業ディレクトリの保持。 |
| ❌ | | **永続化** | `localStorage` や `IndexedDB` への保存・復元。 |
| ❌ | | **排他制御** | `Lock` / `Unlock` ステートメントのエミュレーション。 |
| ⚠️ | **互換性** | **バイナリ/テキスト** | `Binary` / `Random` / `Input` / `Output` 各モードの厳密な挙動。 |

### VFS 開発ロードマップ (TODO)

#### 1. コア機能の強化
- [ ] `MemoryFileSystem` にワイルドカードマッチングロジックの実装（`Kill` への統合）。
- [ ] `MemoryFileSystem` 内部での `cwd` (Current Working Directory) の保持と `ChDir` 対応。
- [ ] `statSync` における VBA 属性（Read-only, Hidden 等）のシミュレーション。

#### 2. ブラウザ環境最適化
- [ ] `FileSystem` インターフェースの非同期版 (`read`, `write` 等) の検討（ブラウザのメインスレッドをブロックしないため）。
- [ ] `IndexedDB` をバックエンドとした `PersistentFileSystem` の実装。

#### 3. テスト環境の改善
- [ ] `filesystem-extra.test.ts` 等のテストコードを、Node.js `fs` 直接参照から `this.fs` (抽象インターフェース) 参照へリファクタリング。
- [ ] 複数の `Evaluator` インスタンス間で単一の `MemoryFileSystem` を共有するためのテスト用設定の追加。

#### 4. バイナリ操作の極致
- [ ] `Put` / `Get` における固定長文字列、UDT、多次元配列のバイナリレイアウトの完全な VBA 互換性検証。

---

## VBA ランタイム挙動

仕様書本文には書かれているが、個別セクションを持たないため見落とされやすい **暗黙のランタイム挙動** の実装状況をトラッキングするセクションです。チェックなしの項目は **未実装 または 未検証** を意味します。実装する際は「ロードマップ」の手順に従い、VBA ベテランが書きそうな代表コードをテストに含めてください。

### オブジェクトのライフサイクルと参照

- [x] **Auto-Instantiation (`Dim x As New ClassName`)**: 遅延インスタンス化 + `Set = Nothing` 後の自動再生成 + `Is Nothing` 常時 False
- [x] **`Dim x As ClassName`（New なし）のデフォルト値**: `Nothing` 初期化
- ✅ **Default Property / Default Member**: `Range("A1") = 10` のような暗黙の `.Value` 解決、Collection の `Item` 暗黙呼び出し等
  - ✅ 包括的なテストスイート作成 (`default-property.test.ts`, 15テスト全パス)
  - ✅ 基盤研究: VBA 仕様書からの解析、実装パターン検討
  - ✅ 実装完了: (1) obj = value -> obj.Value = value (Test 14パス)
  - ✅ 実装完了: (2) obj(args) -> obj.Item(args) (Test 12-13パス)
  - ✅ 実装完了: (3) result = obj -> obj.Value (Test 15パス) - 値コンテキストでの暗黙 Value getter
  - ✅ 修正: Test 7 ByRef パラメータ - implicit Value getter の実装で副次的に修正
- ⚠️ **WithEvents 変数の生存期間**: 親オブジェクト破棄時のイベントハンドラ解除 (制限事項: RaiseEventコア存在; ハンドラ自動登録・検出機構未実装、イベントパラメータ解析未対応) | `withevents-lifecycle.test.ts`
- ✅ **循環参照時の `Set = Nothing` 挙動**: 強制クリアと Class_Terminate の呼び出し順 | テスト: `circular-reference-terminate.test.ts`, `CircularTerminateTest.vba` (VBA: `CircularTestHelper.cls`, `CircularReferenceATest.cls`, `CircularReferenceBTest.cls`, `CircularTerminateTest.vba`)
- ✅ **`Me` キーワードの完全対応**: クラスモジュール内での全コンテキスト | `me-keyword.test.ts`
- [ ] **`Implements` インターフェース呼び出し**: 明示キャスト `Set obj = inst` で interface ビューに切り替わる挙動

### 型変換と暗黙のキャスト

- ✅ **`Variant` 型の値伝播**: `Variant + Number` の結果型の自動拡張 (`variant-propagation.test.ts`)
- [x] **`Empty` / `Null` / `Nothing` / `Missing` の演算子挙動**: `Null + 5 = Null`、`Empty + 5 = 5` などの伝播規則 (`null-empty-propagation.test.ts`)
- [x] **`String` ↔ 数値の暗黙変換**: `"5" + 3 = 8`、`"5" & 3 = "53"` の区別 (`string-numeric-coercion.test.ts`)
- [x] **`Boolean` の代入時の型強制**: 非 0 数値・文字列 "True"/"False"・数値文字列・Empty を Boolean に変換、Null は Error 94 (`boolean-coercion.test.ts`)
- ✅ **Integer/Long のオーバーフロー検出**: Error 6 を正しく発生させる (Byte/Integer/Long/Currency/Single)
- ⚠️ **`Decimal` 型の精度**: 28 桁の固定精度サポート (制限事項: VbaDecimalラッパー存在; 内部はJavaScriptのnumber型でIEEE 754精度、28桁固定小数点精度未実装) | `decimal-precision.test.ts`
- ✅ **`Currency` 型の精度**: 64-bit 整数 / 10000 の固定小数 | `currency-precision.test.ts`

### Variant 配列とコレクション

- ✅ **2 次元 Variant 配列の `For Each` 挙動**: 行優先 vs 列優先の反復順 | `2d-array-foreach.test.ts`
- ✅ **Array() で作る Variant 配列の境界**: `LBound = 0`、明示的 `Option Base` の効果 | `array-lbound-defaults.test.ts`
- ✅ **`ReDim Preserve` の多次元配列での制約**: 最終次元のみ変更可 | `redim-preserve-multidim.test.ts`

### エラーハンドリング

- ✅ **`On Error Resume Next` 下の暗黙の `Err.Clear`**: 正常文実行時のクリア（または非クリア）タイミング | `err-clear-timing.test.ts`
- ✅ **`Resume` の対象決定**: エラー発生点・Resume Next・Resume <label> の正確な制御フロー | `resume-statement-target.test.ts`
- ✅ **エラーハンドラ内での再帰的なエラー発生**: スタックフレームのリセット規則 | `recursive-error-handling.test.ts`
- ⚠️ **`Erl` 関数**: エラー発生行番号の取得（制限事項: MS-VBAL spec未定義、line number tracking未実装） | `erl-function.test.ts`

### プロシージャ呼び出しの細部

- ✅ **モジュール修飾付きプロシージャ呼び出し**: `ModuleName.ProcedureName()` による複数ファイル間での同名関数の区別
  - 実装: `Environment.setProcedureWithModule()` / `getProcedureFromModule()` でモジュール名付き登録・検索
  - VBA 名前解決優先度（修飾なし呼び出し）:
    1. グローバルスコープ（モジュール名なし）← 最優先
    2. 単一モジュール内のプロシージャ（曖昧でない）
    3. 複数モジュール → 曖昧性エラー（修飾必須）
  - 曖昧性検出: 複数モジュールに同名プロシージャがある場合、実行時に詳細エラーを発生
  - テスト: `module-qualified-calls.test.ts`, `ambiguous-procedure-call.test.ts`, `evaluator-scoping.test.ts`
- ✅ **ByRef での文字列・配列・オブジェクトの参照保持**: 文字列・数値・Boolean・配列・オブジェクト全て正常動作 | `byref-reference-preservation.test.ts`
- ⚠️ **ParamArray の境界ケース**: 0 個渡し、配列を 1 つだけ渡したときの展開規則 | `paramarray-edge-cases.test.ts` (制限事項: ByRef semantics未実装)
- ✅ **Optional パラメータの IsMissing 判定**: デフォルト値ありと未指定の区別
- ✅ **Property Get/Let/Set の解決順序**: 同名で混在した場合の優先度 | `property-resolution-order.test.ts`
- ✅ **暗黙の Let（`Call` なしの呼び出し）**: 戻り値が破棄される / されない場面 | `implicit-let.test.ts`

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

- [x] **`CreateObject(progId)` 経由のオブジェクト呼び出し**: 動的ディスパッチ
- ⚠️ **存在しないメソッド呼び出しのエラー詳細**: VBA エラー番号 438 を厳密に再現 (制限事項: late binding error codes未実装)
- ✅ **`IDispatch` 風のプロパティ get/set 自動振り分け**: 引数の有無でメソッド/プロパティ自動判定 | `idispatch-property-dispatch.test.ts`

---

## エンジンの改善

### Parser の拡張機能

- [ ] **Parser に `parseAsClass` パラメータを追加**: `.cls` ファイルやプログラム的にクラスとしてパースすべきコードの指定
  - 現状: VBATest で `.cls` ファイル内のメンバー変数定義を検出して自動的に `Class...End Class` でラップしている
  - 改善案: Parser に `parseAsClass` フラグを追加し、Lexer / Parser で正式にクラスパースモードを切り替える
  - 影響範囲: `src/compiler/parser.ts`, `src/compiler/lexer.ts`, `test-libs/test-runner.ts`
  - 関連: VBA標準では .cls ファイルはファイル名がクラス名となり、本エンジンの拡張機能として Class...End Class 構文を .bas 内でサポート

### VBA 仕様制約の検証

- ✅ **モジュール名の長さ検証（31 文字制限）**: MS-VBAL §5.2 で定義されたモジュール名の最大長を実行時に検証
  - 仕様: モジュール名は最大 31 文字（ファイル名またはVB_Name属性）
  - 実装: `Evaluator.setSourceModule()` でモジュール名長をチェック、超過時にエラー発生
  - テスト: `module-name-length.test.ts`, `module-name-length-integration.test.ts`

- [ ] **静的 vs 動的名前解決の区別が不足**: プロシージャ呼び出し時のエラー検出タイミング
  - 優先度: **中** （Option Explicit あり時は問題なし）
  - 実VBA動作の違い：
    - `unknownProc()` → **コンパイルエラー**（Sub/Function が定義されていません）
    - `unknownModule.unknownProc()` → **実行時エラー 424**（オブジェクトが必要です）
  - 根本原因: VBA仕様では修飾/非修飾で名前解決方式が異なる
    - 非修飾: 静的検証（コンパイル時）
    - 修飾: 動的解決（実行時のメンバアクセス）
  - 当エンジン問題: 両者を同じレベルで処理している（4199-4203行の try-catch が根本原因）
  - 注記: **Option Explicit がない場合の未定義変数の動作**に限定される問題。モダンなVBA開発では Option Explicit が標準であるため、この互換性追求の優先度は低い。Option Explicit ありの場合はコンパイル時に検出される。
  - 修正内容: Parser/Evaluator で修飾/非修飾を区別し、エラーのタイミングを正確に再現
