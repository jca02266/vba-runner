# VBA実装 TODOリスト（MS-VBAL仕様書準拠）

仕様書: `spec/[MS-VBAL]-250520.docx`（テキスト版: `spec/MS-VBAL.txt`）
- **バージョン**: 2.4（2025-05-20）
- **URL**: https://learn.microsoft.com/en-us/openspecs/microsoft_general_purpose_programming_languages/ms-vbal/d5418146-0bd2-45eb-9c7a-fd9502722c74

凡例: ✅ 実装済み / ❌ 未実装 / ⚠️ 部分実装

---

## 第5章：制御ステートメント (§5.4.2)

| 状態 | 機能 | 仕様書 |
|------|------|--------|
| ✅ | Call Statement | §5.4.2.1 |
| ✅ | While Statement (`While...Wend`) | §5.4.2.2 |
| ✅ | For Statement (`For...Next`、Step対応) | §5.4.2.3 |
| ✅ | For Each Statement (`For Each...In`) | §5.4.2.4 |
| ✅ | Exit For Statement | §5.4.2.5 |
| ✅ | Do Statement (`Do While/Until...Loop`, `Do...Loop While/Until`, `Do...Loop`) | §5.4.2.6 |
| ✅ | Exit Do Statement | §5.4.2.7 |
| ✅ | If Statement（複数行） | §5.4.2.8 |
| ✅ | Single-line If Statement | §5.4.2.9 |
| ✅ | Select Case Statement | §5.4.2.10 |
| ✅ | Stop Statement | §5.4.2.11 |
| ✅ | GoTo Statement（汎用ラベルジャンプ） | §5.4.2.12 |
| ✅ | On…GoTo Statement（複数ラベル選択） | §5.4.2.13 |
| ✅ | GoSub Statement | §5.4.2.14 |
| ✅ | Return Statement（GoSub用） | §5.4.2.15 |
| ✅ | On…GoSub Statement | §5.4.2.16 |
| ✅ | Exit Sub Statement | §5.4.2.17 |
| ✅ | Exit Function Statement | §5.4.2.18 |
| ✅ | Exit Property Statement | §5.4.2.19 |
| ❌ | RaiseEvent Statement | §5.4.2.20 |
| ✅ | With Statement (`With...End With`) | §5.4.2.21 |
| ✅ | End Statement（プログラム終了） | §5.4.2.22 |

## 第5章：データ操作ステートメント (§5.4.3)

| 状態 | 機能 | 仕様書 |
|------|------|--------|
| ✅ | Local Variable Declarations (`Dim`) | §5.4.3.1 |
| ✅ | Local Constant Declarations (`Const`) | §5.4.3.2 |
| ⚠️ | ReDim Statement（`Preserve` は簡易実装） | §5.4.3.3 |
| ✅ | Erase Statement | §5.4.3.4 |
| ✅ | Mid Statement（代入形式: `Mid(str, n) = val`） | §5.4.3.5 |
| ✅ | LSet Statement | §5.4.3.6 |
| ✅ | RSet Statement | §5.4.3.7 |
| ✅ | Let Statement（代入） | §5.4.3.8 |
| ✅ | Set Statement（オブジェクト参照） | §5.4.3.9 |

## 第5章：エラーハンドリング (§5.4.4)

| 状態 | 機能 | 仕様書 |
|------|------|--------|
| ✅ | On Error Statement（GoTo / Resume Next / GoTo 0） | §5.4.4.1 |
| ✅ | Resume Statement | §5.4.4.2 |
| ✅ | Error Statement（エラー番号の発生） | §5.4.4.3 |

## 第5章：ファイル操作ステートメント (§5.4.5)

| 状態 | 機能 | 仕様書 |
|------|------|--------|
| ❌ | Open Statement | §5.4.5.1 |
| ❌ | Close / Reset Statements | §5.4.5.2 |
| ❌ | Seek Statement | §5.4.5.3 |
| ❌ | Lock / Unlock Statements | §5.4.5.4/5 |
| ❌ | Line Input Statement | §5.4.5.6 |
| ❌ | Width Statement | §5.4.5.7 |
| ❌ | Print Statement（`Print #n`） | §5.4.5.8 |
| ❌ | Write Statement | §5.4.5.9 |
| ❌ | Input Statement | §5.4.5.10 |
| ❌ | Put Statement | §5.4.5.11 |
| ❌ | Get Statement | §5.4.5.12 |

## 第5章：宣言 (§5.2, §5.3)

| 状態 | 機能 | 仕様書 |
|------|------|--------|
| ✅ | Sub / Function 宣言 | §5.3.1.6 |
| ✅ | Parameter Lists（ByVal, ByRef, Optional） | §5.3.1.5 |
| ✅ | Property Get / Let / Set 宣言 | §5.3.1.7 |
| ❌ | Event Handler 宣言 | §5.3.1.8 |
| ✅ | Variable Declarations（`Dim`） | §5.2.3.1 |
| ✅ | Public / Private / Friend スコープ制御（パースのみ、アクセス制御なし） | §5.2.3.1 / §5.3.1.1 |
| ✅ | Static キーワード（変数・プロシージャ） | §5.3.1.2 |
| ✅ | Const Declarations | §5.2.3.2 |
| ✅ | User Defined Type Declarations（`Type`） | §5.2.3.3 |
| ✅ | Enum Declarations | §5.2.3.4 |
| ❌ | External Procedure Declaration（`Declare`） | §5.2.3.5 |
| ✅ | Class Module（OOP） | §5.2.4 |
| ❌ | Implements Directive | §5.2.4.2 |
| ❌ | Event Declaration | §5.2.4.3 |
| ✅ | Option Explicit（パース済み、実行時は無視） | §5.2 |

## 第5章：演算子 (§5.6.9)

| 状態 | 機能 | 仕様書 |
|------|------|--------|
| ✅ | 算術演算子（`+`, `-`, `*`, `/`, `\`, `Mod`, `^`） | §5.6.9.3 |
| ✅ | 文字列連結（`&`） | §5.6.9.4 |
| ✅ | 比較演算子（`=`, `<>`, `<`, `>`, `<=`, `>=`） | §5.6.9.5 |
| ✅ | Like 演算子 | §5.6.9.6 |
| ✅ | Is 演算子（オブジェクト参照比較） | §5.6.9.7 |
| ✅ | Not 演算子 | §5.6.9.8.1 | (制限事項: Boolean演算時は専用ラッパーを返し、JS側での直接評価には注意が必要) |
| ✅ | And 演算子 | §5.6.9.8.2 | (制限事項: 同上) |
| ✅ | Or 演算子 | §5.6.9.8.3 | (制限事項: 同上) |
| ✅ | Xor 演算子 | §5.6.9.8.4 | (制限事項: 同上) |
| ✅ | Eqv 演算子 | §5.6.9.8.5 | (制限事項: 同上) |
| ✅ | Imp 演算子 | §5.6.9.8.6 | (制限事項: 同上) |
| ✅ | TypeOf…Is 式 | §5.6.7 |
| ✅ | Dictionary Access 式（`!` 演算子） | §5.6.14 |
| ❌ | AddressOf 式 | §5.6.16.8 |

## 第6章：標準ライブラリ — 変換関数 Conversion Module (§6.1.2.3)

| 状態 | 関数 | 仕様書 |
|------|------|--------|
| ✅ | CBool | §6.1.2.3.1.1 | (制限事項: Boolean専用ラッパーを返し、JS側での直接評価には注意が必要) |
| ❌ | CByte | §6.1.2.3.1.2 |
| ❌ | CCur | §6.1.2.3.1.3 |
| ❌ | CDate / CVDate | §6.1.2.3.1.4 |
| ✅ | CDbl | §6.1.2.3.1.5 |
| ❌ | CDec | §6.1.2.3.1.6 |
| ✅ | CInt | §6.1.2.3.1.7 |
| ✅ | CLng | §6.1.2.3.1.8 |
| ❌ | CLngLng | §6.1.2.3.1.9 |
| ❌ | CSng | §6.1.2.3.1.11 |
| ✅ | CStr | §6.1.2.3.1.12 |
| ❌ | CVar | §6.1.2.3.1.13 |
| ❌ | CVErr | §6.1.2.3.1.14 |
| ✅ | Fix | §6.1.2.3.1.16 |
| ✅ | Hex / Hex$ | §6.1.2.3.1.17 |
| ✅ | Int | §6.1.2.3.1.18 |
| ✅ | Oct / Oct$ | §6.1.2.3.1.19 |
| ✅ | Str / Str$ | §6.1.2.3.1.20 |
| ✅ | Val | §6.1.2.3.1.21 |
| ✅ | TypeName | §6.1.2.3.1.22 | (制限事項: 全ての数値はJSのnumber(Double)として扱われ、Integer/Longの区別は未実装) |

## 第6章：標準ライブラリ — 日付/時刻 DateTime Module (§6.1.2.4)

| 状態 | 関数/プロパティ | 仕様書 |
|------|----------------|--------|
| ✅ | DateAdd | §6.1.2.4.1.1 |
| ✅ | DateDiff | §6.1.2.4.1.2 |
| ✅ | DatePart | §6.1.2.4.1.3 |
| ✅ | DateSerial | §6.1.2.4.1.4 |
| ✅ | DateValue | §6.1.2.4.1.5 |
| ✅ | Day | §6.1.2.4.1.6 |
| ✅ | Hour | §6.1.2.4.1.7 |
| ✅ | Minute | §6.1.2.4.1.8 |
| ✅ | Month | §6.1.2.4.1.9 |
| ✅ | Second | §6.1.2.4.1.10 |
| ✅ | TimeSerial | §6.1.2.4.1.11 |
| ✅ | TimeValue | §6.1.2.4.1.12 |
| ✅ | Weekday | §6.1.2.4.1.13 |
| ✅ | Year | §6.1.2.4.1.14 |
| ✅ | Date / Date$ | §6.1.2.4.2.2 |
| ✅ | Now | §6.1.2.4.2.3 |
| ✅ | Time / Time$ | §6.1.2.4.2.4 |
| ✅ | Timer | §6.1.2.4.2.5 |

## 第6章：標準ライブラリ — 数学関数 Math Module (§6.1.2.10)

| 状態 | 関数 | 仕様書 |
|------|------|--------|
| ✅ | Abs | §6.1.2.10.1.1 |
| ✅ | Atn | §6.1.2.10.1.2 |
| ✅ | Cos | §6.1.2.10.1.3 |
| ✅ | Exp | §6.1.2.10.1.4 |
| ✅ | Log | §6.1.2.10.1.5 |
| ✅ | Rnd | §6.1.2.10.1.6 |
| ✅ | Round | §6.1.2.10.1.7 |
| ✅ | Sgn | §6.1.2.10.1.8 |
| ✅ | Sin | §6.1.2.10.1.9 |
| ✅ | Sqr | §6.1.2.10.1.10 |
| ✅ | Tan | §6.1.2.10.1.11 |
| ✅ | Randomize | §6.1.2.10.2.1 |

## 第6章：標準ライブラリ — 文字列関数 Strings Module (§6.1.2.11)

| 状態 | 関数 | 仕様書 |
|------|------|--------|
| ✅ | Asc / AscW | §6.1.2.11.1.1 |
| ✅ | Chr / Chr$ / ChrW | §6.1.2.11.1.4/6 |
| ❌ | Filter | §6.1.2.11.1.7 |
| ❌ | Format / Format$ | §6.1.2.11.1.8/9 |
| ✅ | InStr / InStrB | §6.1.2.11.1.14 |
| ✅ | InStrRev | §6.1.2.11.1.15 |
| ✅ | Join | §6.1.2.11.1.16 |
| ✅ | LCase / LCase$ | §6.1.2.11.1.17/18 |
| ✅ | Left / Left$ | §6.1.2.11.1.19/20 |
| ✅ | Len | §6.1.2.11.1.22 |
| ✅ | LTrim / RTrim | §6.1.2.11.1.23 |
| ✅ | Trim / Trim$ | §6.1.2.11.1.23 |
| ✅ | Mid / Mid$ | §6.1.2.11.1.25/26 |
| ✅ | Replace | §6.1.2.11.1.29 |
| ✅ | Right / Right$ | §6.1.2.11.1.30/31 |
| ✅ | Space / Space$ | §6.1.2.11.1.33/34 |
| ✅ | Split | §6.1.2.11.1.35 |
| ✅ | StrComp | §6.1.2.11.1.36 |
| ❌ | StrConv | §6.1.2.11.1.37 |
| ✅ | String / String$ | §6.1.2.11.1.38/39 |
| ✅ | StrReverse | §6.1.2.11.1.40 |
| ✅ | UCase / UCase$ | §6.1.2.11.1.41/42 |

## 第6章：標準ライブラリ — 情報関数 Information Module (§6.1.2.7)

| 状態 | 関数 | 仕様書 |
|------|------|--------|
| ✅ | IsArray | §6.1.2.7.1.2 |
| ❌ | IsDate | §6.1.2.7.1.3 |
| ✅ | IsEmpty | §6.1.2.7.1.4 |
| ❌ | IsError | §6.1.2.7.1.5 |
| ✅ | IsMissing | §6.1.2.7.1.6 |
| ✅ | IsNull | §6.1.2.7.1.7 |
| ✅ | IsNumeric | §6.1.2.7.1.8 |
| ✅ | IsObject | §6.1.2.7.1.9 |
| ✅ | TypeName | §6.1.2.7.1.12 | (制限事項: 全ての数値はJSのnumber(Double)として扱われ、Integer/Longの区別は未実装) |
| ❌ | VarType | §6.1.2.7.1.13 |

## 第6章：標準ライブラリ — インタラクション Interaction Module (§6.1.2.8)

| 状態 | 関数 | 仕様書 |
|------|------|--------|
| ❌ | Choose | §6.1.2.8.1.2 |
| ⚠️ | CreateObject（`Scripting.Dictionary` のみ対応） | §6.1.2.8.1.4 |
| ❌ | Environ / Environ$ | §6.1.2.8.1.6 |
| ❌ | GetObject | §6.1.2.8.1.9 |
| ✅ | IIf | §6.1.2.8.1.11 |
| ❌ | InputBox | §6.1.2.8.1.12 |
| ❌ | MsgBox | §6.1.2.8.1.13 |
| ❌ | Switch | §6.1.2.8.1.16 |

## 第6章：標準ライブラリ — 配列関連

| 状態 | 関数 | 仕様書 |
|------|------|--------|
| ✅ | UBound | §6（標準ライブラリ） |
| ✅ | LBound | §6（標準ライブラリ） |
| ✅ | Array関数 | §6（標準ライブラリ） |

## 第6章：組み込みクラス (§6.1.3)

### Collection Object (§6.1.3.1)

| 状態 | 機能 | 仕様書 |
|------|------|--------|
| ✅ | Count | §6.1.3.1.1.1 |
| ✅ | Item | §6.1.3.1.1.2 |
| ✅ | Add | §6.1.3.1.2.1 |
| ❌ | Remove | §6.1.3.1.2.2 |

### Err Class (§6.1.3.2)

| 状態 | 機能 | 仕様書 |
|------|------|--------|
| ✅ | Clear | §6.1.3.2.1.1 |
| ✅ | Raise | §6.1.3.2.1.2 |
| ✅ | Description | §6.1.3.2.2.1 |
| ❌ | HelpContext | §6.1.3.2.2.2 |
| ❌ | HelpFile | §6.1.3.2.2.3 |
| ❌ | LastDllError | §6.1.3.2.2.4 |
| ✅ | Number | §6.1.3.2.2.5 |
| ✅ | Source | §6.1.3.2.2.6 |

---

## 実装優先度ガイド

実用上の影響が大きい順に整理した優先実装候補：

### 高優先度（基本制御フロー・頻出構文）

1. **Select Case** (§5.4.2.10) — 条件分岐の基本構文として頻出
2. **For Each...In** (§5.4.2.4) — コレクション・配列の反復に必須
3. **While...Wend** (§5.4.2.2) — 古いVBAコードで頻出
4. **Do...Loop Until** (§5.4.2.6) — ループ末尾条件評価
5. **With...End With** (§5.4.2.21) — オブジェクト操作で多用
6. **GoTo**（汎用） (§5.4.2.12) — エラーハンドラ以外のラベルジャンプ

### 中優先度（標準ライブラリ関数）

7. **文字列関数**: `Len`, `Left`, `Right`, `Mid`, `InStr`, `LCase`, `Split`, `Join`, `Replace`
8. **型変換関数**: `CInt`, `CStr`, `CBool`, `CSng`, `Fix`, `Val`
9. **数学関数**: `Abs`, `Round`, `Sqr`, `Rnd`
10. **情報関数**: `IsArray`, `IsNull`, `IsObject`, `IsMissing`
11. **配列関数**: `LBound`, `Array`
12. **IIf** — インライン条件式として頻用

### 低優先度（特殊用途）

13. `Public` / `Private` スコープ制御
14. `Enum` 宣言
15. `Property Get / Let / Set`
16. 日付/時刻関数（DateTime Module 全般）
17. ファイル操作ステートメント（§5.4.5 全般）
18. `GoSub` / `Return`（レガシー構文）
19. `Like` 演算子
20. 論理演算子: `Xor`, `Eqv`, `Imp`
21. Class Module（OOP全般）
