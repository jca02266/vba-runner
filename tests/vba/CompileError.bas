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
'   VBARunner 対応: VBA の [prerun] は VBARunner では 2 種類に分かれる。
'     resolve : Pass 2（resolveIdentifiers）でモジュールロード時に検出。
'     preproc : Pass 3（precheckProc）で各 Proc 実行直前に AST 静的チェックで検出。
'
' 【自動テスト生成】
'   各 Case_* Sub の直前コメント（CASE/TYPE/VBA/RUNNER）と
'   ボディ内の ' @error マーカーを解析して vba_compile_error.test.ts を生成する。
'   生成コマンド: npx tsx test-libs/compile-error-generator.ts tests/vba/CompileError.bas
'
' 【CASE ブロックのフィールド】
'   CASE    : スネークケースのテスト名
'   TYPE    : parse（VBE で行が赤くなる）、preproc（Pass 3: 各 Proc 実行直前）、または resolve（Pass 2: モジュールロード時）
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
' VBA: コンパイルエラー: 修正候補: 識別子
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
' VBA: コンパイルエラー: 修正候補: 識別子
' RUNNER: /reserved word/i
' NOTE: Sub 宣言でも同様。
'@case-begin
Sub Close() ' @error
End Sub
'@case-end

' CASE: reserved_word_as_function_name_print
' TYPE: parse
' VBA: コンパイルエラー: 修正候補: 識別子
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
' TYPE: resolve
' VBA: コンパイルエラー: 名前が適切ではありません duplicate_sub_name
' RUNNER: /duplicate.*procedure|duplicate.*name/i
' NOTE: 同一モジュール内のプロシージャ名重複（Sub Foo / Sub Foo）
'   resolve だが、このエラーがあると他のエラーの前にこのエラーに引っかかるためコメントアウトしている
'@case-begin
Sub duplicate_sub_name()

End Sub
''Sub duplicate_sub_name()  ' @error
''
''End Sub
'@case-end

' CASE: module_level_dim_after_procedure
' TYPE: resolve
' VBA: コンパイル エラー: End Sub、End Function または End Property 以降には、コメントのみが記述できます。
' RUNNER: /only comments may appear after end sub/i
' EVAL_OPTIONS: { allowTopLevelStatements: false }
' NOTE: 標準 VBA ではモジュールレベルの Dim もすべてのプロシージャより前に書く必要があり、
'   プロシージャ（End Sub 等）の後に書くとコンパイルエラーになる。
'   vba-runner は evalVBASingle/evalVBAModules 用の拡張（allowTopLevelStatements: true が
'   デフォルト）として、先頭にない Dim もモジュールレベル実行文と同様に許容している。
'   このケースは allowTopLevelStatements: false でその拡張を無効化し、標準 VBA 相当の
'   挙動（コンパイルエラー）を検証する。
'@case-begin
Sub ModuleLevelDimAfterProcedure()
End Sub
Dim moduleLevelVar As Integer ' @error
'@case-end

' CASE: module_level_const_after_procedure
' TYPE: resolve
' VBA: コンパイル エラー: End Sub、End Function または End Property 以降には、コメントのみが記述できます。
' RUNNER: /only comments may appear after end sub/i
' EVAL_OPTIONS: { allowTopLevelStatements: false }
' NOTE: モジュールレベルの Const も Dim と同様。allowTopLevelStatements については
'   module_level_dim_after_procedure の NOTE を参照。
'@case-begin
Sub ModuleLevelConstAfterProcedure()
End Sub
Const ModuleLevelConst As Integer = 1 ' @error
'@case-end

' CASE: module_level_public_after_procedure
' TYPE: resolve
' VBA: コンパイル エラー: End Sub、End Function または End Property 以降には、コメントのみが記述できます。
' RUNNER: /only comments may appear after end sub/i
' EVAL_OPTIONS: { allowTopLevelStatements: false }
' NOTE: Public 変数宣言も Dim と同じ VariableDeclaration として扱われ、同様にエラーになる。
'   allowTopLevelStatements については module_level_dim_after_procedure の NOTE を参照。
'@case-begin
Sub ModuleLevelPublicAfterProcedure()
End Sub
Public ModuleLevelPublicVar As Integer ' @error
'@case-end

' CASE: module_level_type_after_procedure
' TYPE: resolve
' VBA: コンパイル エラー: End Sub、End Function または End Property 以降には、コメントのみが記述できます。
' RUNNER: /only comments may appear after end sub/i
' EVAL_OPTIONS: { allowTopLevelStatements: false }
' NOTE: ユーザー定義型（Type ... End Type）も同様。allowTopLevelStatements については
'   module_level_dim_after_procedure の NOTE を参照。
'@case-begin
Sub ModuleLevelTypeAfterProcedure()
End Sub
Type ModuleLevelType ' @error
    Field As Integer
End Type
'@case-end

' CASE: module_level_enum_after_procedure
' TYPE: resolve
' VBA: コンパイル エラー: End Sub、End Function または End Property 以降には、コメントのみが記述できます。
' RUNNER: /only comments may appear after end sub/i
' EVAL_OPTIONS: { allowTopLevelStatements: false }
' NOTE: Enum も同様。allowTopLevelStatements については module_level_dim_after_procedure の
'   NOTE を参照。
'@case-begin
Sub ModuleLevelEnumAfterProcedure()
End Sub
Enum ModuleLevelEnum ' @error
    ModuleLevelEnumValue
End Enum
'@case-end

' CASE: module_level_toplevel_stmt_after_procedure_strict
' TYPE: resolve
' VBA: コンパイル エラー: End Sub、End Function または End Property 以降には、コメントのみが記述できます。
' RUNNER: /only comments may appear after end sub/i
' EVAL_OPTIONS: { allowTopLevelStatements: false }
' NOTE: For/Next・代入などの「モジュールレベル実行文」は標準 VBA では宣言文と同様に扱われ、
'   プロシージャの後に書くと本来コンパイルエラーになる。
'   vba-runner は REPL・テストスクリプト用の拡張として、このような実行文をプロシージャの
'   有無や位置に関係なく許容している（allowTopLevelStatements: true がデフォルト）。
'   このケースは allowTopLevelStatements: false を明示してその拡張を無効化し、
'   標準 VBA 相当の挙動（コンパイルエラー）を検証する。
'   拡張が有効な既定動作（エラーにならないこと）は
'   tests/spec/module-level-toplevel-after-procedure.test.ts で検証する。
'@case-begin
Sub ModuleLevelToplevelStmtAfterProcedureStrict()
End Sub
For moduleLevelIdx = 0 To 10 ' @error
Next moduleLevelIdx
'@case-end

