# VBA実装 TODOリスト（MS-VBAL仕様書準拠）

仕様書: `spec/[MS-VBAL]-250520.docx`（テキスト版: `spec/MS-VBAL.txt`）
- **バージョン**: 2.4（2025-05-20）
- **URL**: https://learn.microsoft.com/en-us/openspecs/microsoft_general_purpose_programming_languages/ms-vbal/d5418146-0bd2-45eb-9c7a-fd9502722c74

凡例: ✅ 実装済み / ❌ 未実装 / ⚠️ 部分実装

---

## 第5章：制御ステートメント (§5.4.2)

| 状態 | 優先度 | 機能 | 仕様書 |
|------|--------|------|--------|
| ✅ | P0 | Call Statement | §5.4.2.1 |
| ✅ | P0 | While Statement (`While...Wend`) | §5.4.2.2 |
| ✅ | P0 | For Statement (`For...Next`、Step対応) | §5.4.2.3 |
| ✅ | P0 | For Each Statement (`For Each...In`) | §5.4.2.4 |
| ✅ | P1 | Exit For Statement | §5.4.2.5 |
| ✅ | P0 | Do Statement (`Do While/Until...Loop`, `Do...Loop While/Until`, `Do...Loop`) | §5.4.2.6 |
| ✅ | P1 | Exit Do Statement | §5.4.2.7 |
| ✅ | P0 | If Statement（複数行） | §5.4.2.8 |
| ✅ | P0 | Single-line If Statement | §5.4.2.9 |
| ✅ | P0 | Select Case Statement | §5.4.2.10 |
| ✅ | P1 | Stop Statement | §5.4.2.11 |
| ✅ | P1 | GoTo Statement（汎用ラベルジャンプ） | §5.4.2.12 |
| ✅ | P0 | On…GoTo Statement（複数ラベル選択） | §5.4.2.13 |
| ✅ | P2 | GoSub Statement | §5.4.2.14 |
| ✅ | P2 | Return Statement（GoSub用） | §5.4.2.15 |
| ✅ | P2 | On…GoSub Statement | §5.4.2.16 |
| ✅ | P1 | Exit Sub Statement | §5.4.2.17 |
| ✅ | P1 | Exit Function Statement | §5.4.2.18 |
| ✅ | P1 | Exit Property Statement | §5.4.2.19 |
| ❌ | P2 | RaiseEvent Statement | §5.4.2.20 |
| ✅ | P0 | With Statement (`With...End With`) | §5.4.2.21 |
| ✅ | P1 | End Statement（プログラム終了） | §5.4.2.22 |
| ✅ | P1 | Assert Statement | §5.4.2.23 |

## 第5章：データ操作ステートメント (§5.4.3)

| 状態 | 機能 | 仕様書 |
|------|------|--------|
| ✅ | P0 | Local Variable Declarations (`Dim`) | §5.4.3.1 |
| ✅ | P0 | Local Constant Declarations (`Const`) | §5.4.3.2 |
| ⚠️ | P1 | ReDim Statement（`Preserve` は簡易実装） | §5.4.3.3 |
| ✅ | P1 | Erase Statement | §5.4.3.4 |
| ✅ | P1 | Mid Statement（代入形式: `Mid(str, n) = val`） | §5.4.3.5 |
| ❌ | P2 | MidB / Mid$ / MidB$ Statement | §5.4.3.5 |
| ✅ | P1 | LSet Statement | §5.4.3.6 |
| ✅ | P1 | RSet Statement | §5.4.3.7 |
| ✅ | P0 | Let Statement（代入） | §5.4.3.8 |
| ✅ | P0 | Set Statement（オブジェクト参照） | §5.4.3.9 |

## 第5章：エラーハンドリング (§5.4.4)

| 状態 | 優先度 | 機能 | 仕様書 |
|------|--------|------|--------|
| ✅ | P0 | On Error Statement（GoTo / Resume Next / GoTo 0） | §5.4.4.1 |
| ✅ | P0 | Resume Statement | §5.4.4.2 |
| ✅ | P2 | Error Statement（エラー番号の発生） | §5.4.4.3 |

## 第5章：ファイル操作ステートメント (§5.4.5)

| 状態 | 優先度 | 機能 | 仕様書 |
|------|--------|------|--------|
| ❌ | P1 | Open Statement | §5.4.5.1 |
| ❌ | P1 | Close / Reset Statements | §5.4.5.2 |
| ❌ | P2 | Seek Statement | §5.4.5.3 |
| ❌ | P2 | Lock / Unlock Statements | §5.4.5.4/5 |
| ❌ | P2 | Line Input Statement | §5.4.5.6 |
| ❌ | P2 | Width Statement | §5.4.5.7 |
| ❌ | P2 | Print Statement（`Print #n`） | §5.4.5.8 |
| ❌ | P2 | Write Statement | §5.4.5.9 |
| ❌ | P2 | Input Statement | §5.4.5.10 |
| ❌ | P2 | Put Statement | §5.4.5.11 |
| ❌ | P2 | Get Statement | §5.4.5.12 |

## 第5章：宣言 (§5.2, §5.3)

| 状態 | 優先度 | 機能 | 仕様書 |
|------|--------|------|--------|
| ✅ | P0 | Sub / Function 宣言 | §5.3.1.6 |
| ✅ | P0 | Parameter Lists（ByVal, ByRef, Optional） | §5.3.1.5 |
| ✅ | P0 | Property Get / Let / Set 宣言 | §5.3.1.7 |
| ❌ | P2 | Event Handler 宣言 | §5.3.1.8 |
| ✅ | P0 | Variable Declarations（`Dim`） | §5.2.3.1 |
| ✅ | P1 | Public / Private / Friend スコープ制御（パースのみ、アクセス制御なし） | §5.2.3.1 / §5.3.1.1 |
| ✅ | P1 | Static キーワード（変数・プロシージャ） | §5.3.1.2 |
| ✅ | P0 | Const Declarations | §5.2.3.2 |
| ✅ | P1 | User Defined Type Declarations（`Type`） | §5.2.3.3 |
| ✅ | P1 | Enum Declarations | §5.2.3.4 |
| ✅ | P0 | External Procedure Declaration（`Declare`） | §5.2.3.5 | (制限事項: JSランタイム上ではスタブとして登録) |
| ✅ | P0 | Class Module（OOP） | §5.2.4 |
| ❌ | P1 | Implements Directive | §5.2.4.2 |
| ❌ | P2 | Event Declaration | §5.2.4.3 |
| ✅ | P1 | Option Explicit（パース済み、実行時は無視） | §5.2 |
| ✅ | P0 | Option Compare | §5.2.1.1 |
| ✅ | P1 | Option Base | §5.2.1.2 |
| ✅ | P1 | Option Private Module | §5.2.1.4 |
| ✅ | P0 | Attributes (VB_Name, etc.) | §5.2.3.1.6 / §5.2.4.1 | (制限事項: パースのみ。実行時は無視) |
| ✅ | P0 | Date Literals (#mm/dd/yyyy#) | §3.3.3.3 |

## 第5章：演算子 (§5.6.9)

| 状態 | 優先度 | 機能 | 仕様書 |
|------|--------|------|--------|
| ✅ | P1 | 算術演算子（`+`, `-`, `*`, `/`, `\`, `Mod`, `^`） | §5.6.9.3 |
| ✅ | P1 | 文字列連結（`&`） | §5.6.9.4 |
| ✅ | P1 | 比較演算子（`=`, `<>`, `<`, `>`, `<=`, `>=`） | §5.6.9.5 |
| ✅ | P1 | Like 演算子 | §5.6.9.6 |
| ✅ | P1 | Is 演算子（オブジェクト参照比較） | §5.6.9.7 |
| ✅ | P1 | Not 演算子 | §5.6.9.8.1 | (制限事項: Boolean演算時は専用ラッパーを返し、JS側での直接評価には注意が必要) |
| ✅ | P1 | And 演算子 | §5.6.9.8.2 | (制限事項: 同上) |
| ✅ | P1 | Or 演算子 | §5.6.9.8.3 | (制限事項: 同上) |
| ✅ | P1 | Xor 演算子 | §5.6.9.8.4 | (制限事項: 同上) |
| ✅ | P1 | Eqv 演算子 | §5.6.9.8.5 | (制限事項: 同上) |
| ✅ | P1 | Imp 演算子 | §5.6.9.8.6 | (制限事項: 同上) |
| ✅ | P1 | TypeOf…Is 式 | §5.6.7 |
| ✅ | P1 | Dictionary Access 式（`!` 演算子） | §5.6.14 |
| ❌ | P2 | AddressOf 式 | §5.6.16.8 |

## 第6章：標準ライブラリ — 変換関数 Conversion Module (§6.1.2.3)

| 状態 | 優先度 | 関数 | 仕様書 |
|------|--------|------|--------|
| ✅ | P1 | CBool | §6.1.2.3.1.1 | (制限事項: Boolean専用ラッパーを返し、JS側での直接評価には注意が必要) |
| ✅ | P1 | CByte | §6.1.2.3.1.2 |
| ✅ | P2 | CCur | §6.1.2.3.1.3 |
| ✅ | P0 | CDate / CVDate | §6.1.2.3.1.4 |
| ✅ | P1 | CDbl | §6.1.2.3.1.5 |
| ❌ | P2 | CDec | §6.1.2.3.1.6 |
| ✅ | P1 | CInt | §6.1.2.3.1.7 |
| ✅ | P1 | CLng | §6.1.2.3.1.8 |
| ❌ | P2 | CLngLng | §6.1.2.3.1.9 |
| ✅ | P2 | CSng | §6.1.2.3.1.11 |
| ✅ | P1 | CStr | §6.1.2.3.1.12 |
| ❌ | P2 | CVar | §6.1.2.3.1.13 |
| ❌ | P2 | CVErr | §6.1.2.3.1.14 |
| ❌ | P2 | Error / Error$ | §6.1.2.3.1.15 |
| ✅ | P1 | Fix | §6.1.2.3.1.16 |
| ✅ | P1 | Hex / Hex$ | §6.1.2.3.1.17 |
| ✅ | P1 | Int | §6.1.2.3.1.18 |
| ✅ | P1 | Oct / Oct$ | §6.1.2.3.1.19 |
| ✅ | P1 | Str / Str$ | §6.1.2.3.1.20 |
| ✅ | P1 | Val | §6.1.2.3.1.21 |
| ✅ | P1 | TypeName | §6.1.2.3.1.22 | (制限事項: 全ての数値はJSのnumber(Double)として扱われ、Integer/Longの区別は未実装) |

## 第6章：標準ライブラリ — 日付/時刻 DateTime Module (§6.1.2.4)

| 状態 | 優先度 | 関数/プロパティ | 仕様書 |
|------|--------|----------------|--------|
| ✅ | P1 | DateAdd | §6.1.2.4.1.1 |
| ✅ | P1 | DateDiff | §6.1.2.4.1.2 |
| ✅ | P1 | DatePart | §6.1.2.4.1.3 |
| ✅ | P0 | DateSerial | §6.1.2.4.1.4 |
| ✅ | P1 | DateValue | §6.1.2.4.1.5 |
| ✅ | P1 | Day | §6.1.2.4.1.6 |
| ✅ | P1 | Hour | §6.1.2.4.1.7 |
| ✅ | P1 | Minute | §6.1.2.4.1.8 |
| ✅ | P1 | Month | §6.1.2.4.1.9 |
| ✅ | P1 | Second | §6.1.2.4.1.10 |
| ✅ | P1 | TimeSerial | §6.1.2.4.1.11 |
| ✅ | P1 | TimeValue | §6.1.2.4.1.12 |
| ✅ | P1 | Weekday | §6.1.2.4.1.13 |
| ✅ | P1 | Year | §6.1.2.4.1.14 |
| ✅ | P1 | Date / Date$ | §6.1.2.4.2.2 |
| ✅ | P1 | Now | §6.1.2.4.2.3 |
| ✅ | P1 | Time / Time$ | §6.1.2.4.2.4 |
| ✅ | P1 | Timer | §6.1.2.4.2.5 |

## 第6章：標準ライブラリ — 数学関数 Math Module (§6.1.2.10)

| 状態 | 優先度 | 関数 | 仕様書 |
|------|--------|------|--------|
| ✅ | P1 | Abs | §6.1.2.10.1.1 |
| ✅ | P2 | Atn | §6.1.2.10.1.2 |
| ✅ | P1 | Cos | §6.1.2.10.1.3 |
| ✅ | P2 | Exp | §6.1.2.10.1.4 |
| ✅ | P2 | Log | §6.1.2.10.1.5 |
| ✅ | P1 | Rnd | §6.1.2.10.1.6 |
| ✅ | P1 | Round | §6.1.2.10.1.7 |
| ✅ | P2 | Sgn | §6.1.2.10.1.8 |
| ✅ | P1 | Sin | §6.1.2.10.1.9 |
| ✅ | P1 | Sqr | §6.1.2.10.1.10 |
| ✅ | P1 | Tan | §6.1.2.10.1.11 |
| ✅ | P1 | Randomize | §6.1.2.10.2.1 |

## 第6章：標準ライブラリ — 文字列関数 Strings Module (§6.1.2.11)

| 状態 | 優先度 | 関数 | 仕様書 |
|------|--------|------|--------|
| ✅ | P1 | Asc / AscW | §6.1.2.11.1.1 |
| ✅ | P1 | Chr / Chr$ / ChrW | §6.1.2.11.1.4/6 |
| ✅ | P1 | Filter | §6.1.2.11.1.7 |
| ✅ | P0 | Format / Format$ | §6.1.2.11.1.8/9 |
| ✅ | P0 | InStr / InStrB | §6.1.2.11.1.14 |
| ✅ | P1 | InStrRev | §6.1.2.11.1.15 |
| ✅ | P1 | Join | §6.1.2.11.1.16 |
| ✅ | P1 | LCase / LCase$ | §6.1.2.11.1.17/18 |
| ✅ | P1 | Left / Left$ | §6.1.2.11.1.19/20 |
| ✅ | P0 | Len | §6.1.2.11.1.22 |
| ✅ | P1 | LTrim / RTrim | §6.1.2.11.1.23 |
| ✅ | P1 | Trim / Trim$ | §6.1.2.11.1.23 |
| ✅ | P0 | Mid / Mid$ | §6.1.2.11.1.25/26 |
| ✅ | P1 | Replace | §6.1.2.11.1.29 |
| ✅ | P1 | Right / Right$ | §6.1.2.11.1.30/31 |
| ✅ | P1 | Space / Space$ | §6.1.2.11.1.33/34 |
| ✅ | P1 | Split | §6.1.2.11.1.35 |
| ✅ | P1 | StrComp | §6.1.2.11.1.36 |
| ✅ | P2 | StrConv | §6.1.2.11.1.37 |
| ✅ | P1 | String / String$ | §6.1.2.11.1.38/39 |
| ✅ | P1 | StrReverse | §6.1.2.11.1.40 |
| ✅ | P1 | UCase / UCase$ | §6.1.2.11.1.41/42 |

## 第6章：標準ライブラリ — 情報関数 Information Module (§6.1.2.7)

| 状態 | 優先度 | 関数 | 仕様書 |
|------|--------|------|--------|
| ✅ | P1 | IsArray | §6.1.2.7.1.2 |
| ✅ | P0 | IsDate | §6.1.2.7.1.3 |
| ✅ | P1 | IsEmpty | §6.1.2.7.1.4 |
| ✅ | P1 | IsError | §6.1.2.7.1.5 |
| ✅ | P1 | IsMissing | §6.1.2.7.1.6 |
| ✅ | P1 | IsNull | §6.1.2.7.1.7 |
| ✅ | P1 | IsNumeric | §6.1.2.7.1.8 |
| ✅ | P1 | IsObject | §6.1.2.7.1.9 |
| ✅ | P0 | TypeName | §6.1.2.7.1.12 | (制限事項: 全ての数値はJSのnumber(Double)として扱われ、Integer/Longの区別は未実装) |
| ✅ | P0 | VarType | §6.1.2.7.1.13 |

## 第6章：標準ライブラリ — インタラクション Interaction Module (§6.1.2.8)

| 状態 | 優先度 | 関数 | 仕様書 |
|------|--------|------|--------|
| ✅ | P1 | Choose | §6.1.2.8.1.2 |
| ✅ | P0 | CreateObject (拡張) | §6.1.2.8.1.4 | (制限事項: Scripting.Dictionary等の主要オブジェクトのみ対応) |
| ❌ | P2 | Environ / Environ$ | §6.1.2.8.1.6 |
| ❌ | P1 | GetObject | §6.1.2.8.1.9 |
| ✅ | P1 | IIf | §6.1.2.8.1.11 |
| ✅ | P0 | InputBox | §6.1.2.8.1.12 | (制限事項: ダミー実装。コンソール出力と固定値返却) |
| ✅ | P0 | MsgBox | §6.1.2.8.1.13 | (制限事項: ダミー実装。コンソール出力と固定値返却) |
| ✅ | P1 | Switch | §6.1.2.8.1.16 |

## 第6章：標準ライブラリ — 配列関連

| 状態 | 優先度 | 関数 | 仕様書 |
|------|--------|------|--------|
| ✅ | P0 | UBound | §6（標準ライブラリ） |
| ✅ | P0 | LBound | §6（標準ライブラリ） |
| ✅ | P0 | Array関数 | §6（標準ライブラリ） |

## 第6章：標準ライブラリ — ファイルシステム FileSystem Module (§6.1.2.5)

| 状態 | 優先度 | 関数/サブプロシージャ | 仕様書 |
|------|--------|----------------|--------|
| ❌ | P2 | CurDir / CurDir$ | §6.1.2.5.1.1 |
| ❌ | P2 | Dir | §6.1.2.5.1.2 |
| ❌ | P2 | EOF | §6.1.2.5.1.3 |
| ❌ | P2 | FileAttr | §6.1.2.5.1.4 |
| ❌ | P2 | FileDateTime | §6.1.2.5.1.5 |
| ❌ | P2 | FileLen | §6.1.2.5.1.6 |
| ❌ | P1 | FreeFile | §6.1.2.5.1.7 |
| ❌ | P2 | Loc | §6.1.2.5.1.8 |
| ❌ | P2 | LOF | §6.1.2.5.1.9 |
| ❌ | P2 | Seek | §6.1.2.5.1.10 |
| ❌ | P2 | ChDir | §6.1.2.5.2.1 |
| ❌ | P2 | ChDrive | §6.1.2.5.2.2 |
| ❌ | P2 | FileCopy | §6.1.2.5.2.3 |
| ❌ | P2 | Kill | §6.1.2.5.2.4 |
| ❌ | P2 | MkDir | §6.1.2.5.2.5 |
| ❌ | P2 | RmDir | §6.1.2.5.2.6 |
| ❌ | P2 | SetAttr | §6.1.2.5.2.7 |

## 第6章：標準ライブラリ — 財務 Financial Module (§6.1.2.6)

| 状態 | 優先度 | 関数 | 仕様書 |
|------|--------|------|--------|
| ❌ | P2 | DDB | §6.1.2.6.1.1 |
| ❌ | P2 | FV | §6.1.2.6.1.2 |
| ❌ | P2 | IPmt | §6.1.2.6.1.3 |
| ❌ | P2 | IRR | §6.1.2.6.1.4 |
| ❌ | P2 | MIRR | §6.1.2.6.1.5 |
| ❌ | P2 | NPer | §6.1.2.6.1.6 |
| ❌ | P2 | NPV | §6.1.2.6.1.7 |
| ❌ | P2 | Pmt | §6.1.2.6.1.8 |
| ❌ | P2 | PPmt | §6.1.2.6.1.9 |
| ❌ | P2 | PV | §6.1.2.6.1.10 |
| ❌ | P2 | Rate | §6.1.2.6.1.11 |
| ❌ | P2 | SLN | §6.1.2.6.1.12 |
| ❌ | P2 | SYD | §6.1.2.6.1.13 |

## 第6章：組み込みクラス (§6.1.3)

### Collection Object (§6.1.3.1)

| 状態 | 優先度 | 機能 | 仕様書 |
|------|--------|------|--------|
| ✅ | P1 | Count | §6.1.3.1.1.1 |
| ✅ | P1 | Item | §6.1.3.1.1.2 |
| ✅ | P1 | Add | §6.1.3.1.2.1 |
| ✅ | P1 | Remove | §6.1.3.1.2.2 |

### Err Class (§6.1.3.2)

| 状態 | 優先度 | 機能 | 仕様書 |
|------|--------|------|--------|
| ✅ | P1 | Clear | §6.1.3.2.1.1 |
| ✅ | P1 | Raise | §6.1.3.2.1.2 |
| ✅ | P1 | Description | §6.1.3.2.2.1 |
| ❌ | P2 | HelpContext | §6.1.3.2.2.2 |
| ❌ | P2 | HelpFile | §6.1.3.2.2.3 |
| ❌ | P2 | LastDllError | §6.1.3.2.2.4 |
| ✅ | P1 | Number | §6.1.3.2.2.5 |
| ✅ | P1 | Source | §6.1.3.2.2.6 |

---

## 実装優先度ガイド

実用上の影響が大きい順に整理した優先実装候補：

### P0：最優先（基本制御フロー・頻出構文・標準関数）

1.  **CDate / CVDate** (§6.1.2.3.1.4) — 日付リテラル活用に必須
2.  **Format / Format$** (§6.1.2.11.1.8) — 出力整形に多用
3.  **MsgBox / InputBox** (§6.1.2.11.1.x) — 対話型スクリプトの維持
4.  **Option Compare** (§5.2.1.1) — 文字列比較の正確性
5.  **Attributes** (§5.2.3.1.6) — ソースエクスポートファイルの直接読込
6.  **CreateObject (拡張)** (§6.1.2.8.1.4) — 外部オブジェクト連携の強化
7.  **エラーハンドリング** (§5.4.4) — `On Error GoTo`, `Resume` の完全な準拠

### P1：重要（標準ライブラリの拡充）

7.  **文字列関数**: `Filter`, `StrConv`
8.  **変換関数**: `CByte`, `CCur`, `CDec`, `CSng`
9.  **数学関数**: `Abs`, `Round` 等の未実装エッジケース
10. **制御文**: `Assert` ステートメント

### P2：低優先度（特殊用途・レガシー）

11. 財務関数 (Financial Module) 全般
12. ファイル操作ステートメント (Open/Print等) — Node.js環境での設計が必要
13. レガシー構文: `GoSub` 等
