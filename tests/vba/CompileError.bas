Option Explicit
' ================================================================
' CompileError.bas - VBA コンパイルエラー仕様記述ファイル
'
' 【VBA コンパイルエラーの 2 種類】
'
' [parse] 構文エラー（実行前）:
'   VBE で該当行が赤くなり、実行ボタンを押す前からエラーがわかる。
'   → このファイルを VBE に貼り付けると該当行がすぐ赤くなる。
'   VBARunner 対応: Parser.parse() 時に例外が発生（パース時エラー）。
'
' [prerun] 実行直前コンパイルエラー:
'   実行ボタンを押すと VBE カーソルが該当箇所に移動し、実行直前に
'   コンパイルエラーダイアログが表示される。On Error では捕捉できない。
'   → [parse] の Sub をコメントアウトした上で各 Sub にカーソルを置き F5 で確認。
'   VBARunner 対応: プロシージャ呼び出し直前の静的チェックで例外が発生。
'
' 【自動テスト生成】
'   各 Case_* Sub の直前コメント（CASE/TYPE/VBA/RUNNER）と
'   ボディ内の ' @error マーカーを解析して vba_compile_error.test.ts を生成する。
'   生成コマンド: npx tsx test-libs/compile-error-generator.ts tests/vba/CompileError.bas
'
' 【CASE ブロックのフィールド】
'   CASE    : スネークケースのテスト名
'   TYPE    : parse（VBE で行が赤くなる）または prerun（実行直前に検出）
'   VBA     : 実 VBA が表示するエラーメッセージ（仕様。変更不可）
'   RUNNER  : VBARunner が投げる例外の正規表現（VBA の意味を包含し拡張可）
'   NOTE    : 実装状況などの補足（省略可）
'
' 【エラー行マーカー】
'   Sub ボディ内のエラー発生行末尾に ' @error を付ける。
'   ジェネレーターはこのマーカーから行番号を自動計算し、
'   テストコードからはマーカーを除去して出力する。
' ================================================================

' ----------------------------------------------------------------
' ヘルパー定義（各テストケースが参照するプロシージャ）
' ----------------------------------------------------------------
Private Sub MySub()
End Sub

Private Function MyFuncHasArg(x)
End Function

' ================================================================
' [parse] 構文エラー
' VBE に貼り付けると該当行が赤くなることで確認できる。
' ※ [prerun] の Sub を実行する際はこのセクションをコメントアウトすること。
' ================================================================

' CASE: sub_call_with_empty_parens
' TYPE: parse
' VBA: コンパイルエラー: 構文エラー
' RUNNER: /syntax error|parse error/i
Sub Case_sub_call_with_empty_parens1()
    MySub() ' @error
End Sub
Sub Case_sub_call_with_empty_parens2()
    MyFunction() ' @error
End Sub
Sub Case_sub_call_with_empty_parens3()
    MySub () ' @error
End Sub
Sub Case_sub_call_with_empty_parens4()
    UnKnown() ' @error
End Sub

' CASE: sub_call_arg_without_call_keyword
' TYPE: parse
' VBA: コンパイルエラー: 構文エラー
' RUNNER: /syntax error|parse error/i
Sub Case_sub_call_arg_without_call_keyword()
    Call MySub 42 ' @error
End Sub

' CASE: assign_func_arg_no_parens
' TYPE: parse
' VBA: コンパイルエラー: 構文エラー
' RUNNER: /syntax error|parse error/i
Sub Case_assign_func_arg_no_parens()
    Dim v
    v = MyFuncHasArg arg ' @error
End Sub

' CASE: label_then_sub_call_with_empty_parens
' TYPE: parse
' VBA: コンパイルエラー: 構文エラー
' RUNNER: /syntax error|parse error/i
' NOTE: MySub: MySub はエラーにならない（前者ラベル定義、後者 Sub 呼び出し）。
'       MySub: MySub() は MySub() 部分が空括弧で構文エラー。
Sub Case_label_then_sub_call_with_empty_parens1()
    MySub: MySub() ' @error
End Sub
Sub Case_label_then_sub_call_with_empty_parens2()
    MySub : MySub() ' @error
End Sub

' CASE: reserved_word_as_function_name
' TYPE: parse
' VBA: コンパイルエラー: 予約語はプロシージャ名として使用できません
' RUNNER: /reserved word/i
' NOTE: statement-keyword (Open/Close/Print/Input/Write/Put/Seek/Lock/Unlock 等) は
'   §3.3.5.2 で reserved-identifier のためモジュールレベルのプロシージャ名として使用不可。
'   実際の VBE はこれらを許容する場合があるが、仕様準拠の実装ではコンパイルエラー。
'@case-begin
Function Open() As Long ' @error
    Open = 1
End Function
'@case-end

' CASE: reserved_word_as_sub_name
' TYPE: parse
' VBA: コンパイルエラー: 予約語はプロシージャ名として使用できません
' RUNNER: /reserved word/i
' NOTE: Sub 宣言でも同様。
'@case-begin
Sub Close() ' @error
End Sub
'@case-end

' CASE: reserved_word_as_function_name_print
' TYPE: parse
' VBA: コンパイルエラー: 予約語はプロシージャ名として使用できません
' RUNNER: /reserved word/i
'@case-begin
Function Print() As Long ' @error
    Print = 99
End Function
'@case-end

' ================================================================
' [prerun] 実行直前コンパイルエラー
' 上記 [parse] の Sub をコメントアウトした状態で、
' 各 Sub にカーソルを置いて F5 で実行するとエラーが確認できる。
' ================================================================

' CASE: assign_from_sub
' TYPE: preproc
' VBA: コンパイルエラー: FunctionまたはVariableが必要です
' RUNNER: /function or variable/i
Sub Case_assign_from_sub()
    Dim v
    v = MySub ' @error
End Sub

' CASE: assign_from_sub_with_parens
' TYPE: preproc
' VBA: コンパイルエラー: FunctionまたはVariableが必要です
' RUNNER: /function or variable/i
Sub Case_assign_from_sub_with_parens()
    Dim v
    v = MySub() ' @error
End Sub

' CASE: duplicate_dim
' TYPE: preproc
' VBA: コンパイルエラー: 同じ適用範囲内で宣言が重複しています
' RUNNER: /duplicate/i
Sub Case_duplicate_dim()
    Dim v
    Dim v ' @error
End Sub

' CASE: goto_undefined_label
' TYPE: preproc
' VBA: コンパイルエラー: 行ラベルが定義されていません
' RUNNER: /not defined.*label|label.*not defined/i
Sub Case_goto_undefined_label()
    GoTo NoSuchLabel ' @error
End Sub

' ----------------------------------------------------------------
' 備考（コンパイルエラーにならない例）:
'
'   MySub: MySub     → エラーにならない。前者はラベル定義、後者は Sub 呼び出し。
'   MySub : MySub    → 同上（コロン前後のスペースは無関係）。
'
'   MySub(42) は VBE 上では自動フォーマットで `MySub (42)` に変換され、
'   コンパイルエラーにはならない（42 を括弧式として評価し Sub に渡す VBA 仕様）。
'
'   UnknownModule.UnknownProc（Option Explicit なし）→ 実行時エラー 424（Object required）。
'     Option Explicit あり の場合は変数未宣言コンパイルエラー（下記 Case_qualified_undeclared 参照）。
'
' 未実装（VBARunner が未検出）:
'   同一モジュール内のプロシージャ名重複（Sub Foo / Sub Foo）→ VBE では prerun エラー
' ----------------------------------------------------------------
' CASE: undefined_sub_call
' TYPE: preproc
' VBA: コンパイルエラー: SubまたはFunctionが定義されていません
' RUNNER: /sub or function not defined/i
' NOTE: 非修飾の未定義プロシージャ呼び出しは静的検証（コンパイル時）でエラー。
'   修飾付き（UnknownModule.UnknownProc）は実行時エラー 424（動的解決）のためここに含めない。
Sub Case_undefined_sub_call()
    UnknownProc ' @error
End Sub

' CASE: qualified_undeclared_obj
' TYPE: preproc
' VBA: コンパイルエラー: 変数が定義されていません（Option Explicit で unknownModule が未宣言）
' RUNNER: /variable not declared|not declared/i
' NOTE: Option Explicit 有効時、修飾付き呼び出し UnknownModule.Method の
'   オブジェクト部分 UnknownModule が未宣言変数として検出される。
'   Option Explicit がない場合は Dim が暗黙挿入されて実行時エラー 424 になる。
'@case-begin
Option Explicit
Sub Case_qualified_undeclared_obj()
    UnknownModule.UnknownProc ' @error
End Sub
'@case-end

' CASE: undefined_sub_call_no_oe
' TYPE: preproc
' VBA: コンパイルエラー: SubまたはFunctionが定義されていません
' RUNNER: /sub or function not defined/i
' NOTE: Option Explicit なし。未定義プロシージャは OE の有無に関わらず prerun エラー。
'@case-begin
Sub Case_undefined_sub_call_no_oe()
    UnknownProc ' @error
End Sub
'@case-end

' CASE: undefined_sub_call_with_oe
' TYPE: preproc
' VBA: コンパイルエラー: SubまたはFunctionが定義されていません
' RUNNER: /sub or function not defined/i
' NOTE: Option Explicit あり。未定義プロシージャは OE 違反ではなく Sub or Function not defined エラー。
'@case-begin
Option Explicit
Sub Case_undefined_sub_call_with_oe()
    UnknownProc ' @error
End Sub
'@case-end

' CASE: sub_call_arg_count_mismatch
' TYPE: preproc
' VBA: コンパイルエラー: 引数の数が一致していません。または不正なプロパティを指定しています。
' RUNNER: /wrong number of arguments/i
' NOTE: MySub(42) は VBE 上では自動フォーマットで `MySub (42)` に変換され
'   コンパイルエラーにはならない（42 を括弧式として評価し Sub に渡す VBA 仕様）。
Sub Case_sub_call_arg_count_mismatch()
    MySub (42)  ' @error
End Sub

' CASE: duplicate_sub_name
' TYPE: prerun
' VBA: コンパイルエラー: 名前が適切ではありません duplicate_sub_name
' RUNNER: /duplicate.*procedure|duplicate.*name/i
' NOTE: 同一モジュール内のプロシージャ名重複（Sub Foo / Sub Foo）
'   prerun だが、このエラーがあると他のエラーの前にこのエラーに引っかかるためコメントアウトしている
'@case-begin
Sub duplicate_sub_name()

End Sub
''Sub duplicate_sub_name()  ' @error
''
''End Sub
'@case-end

