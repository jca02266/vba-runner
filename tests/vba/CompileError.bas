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
'   各 Case_* Sub の直前コメント（CASE/TYPE/VBA/RUNNER）を解析して
'   vba_compile_error.test.ts を生成する。
'   生成コマンド: npx tsx test-libs/compile-error-generator.ts tests/vba/CompileError.bas
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
' RUNNER: /syntax error/i
Sub Case_sub_call_with_empty_parens()
    MySub()
End Sub

' CASE: sub_call_arg_without_call_keyword
' TYPE: parse
' VBA: コンパイルエラー: 構文エラー
' RUNNER: /syntax error/i
Sub Case_sub_call_arg_without_call_keyword()
    Call MySub 42
End Sub

' CASE: assign_func_arg_no_parens
' TYPE: parse
' VBA: コンパイルエラー: 構文エラー
' RUNNER: /syntax error/i
Sub Case_assign_func_arg_no_parens()
    Dim v
    v = MyFuncHasArg arg
End Sub

' ================================================================
' [prerun] 実行直前コンパイルエラー
' 上記 [parse] の Sub をコメントアウトした状態で、
' 各 Sub にカーソルを置いて F5 で実行するとエラーが確認できる。
' ================================================================

' CASE: assign_from_sub
' TYPE: prerun
' VBA: コンパイルエラー: FunctionまたはVariableが必要です
' RUNNER: /function or variable/i
Sub Case_assign_from_sub()
    Dim v
    v = MySub
End Sub

' CASE: assign_from_sub_with_parens
' TYPE: prerun
' VBA: コンパイルエラー: FunctionまたはVariableが必要です
' RUNNER: /function or variable/i
Sub Case_assign_from_sub_with_parens()
    Dim v
    v = MySub()
End Sub

' CASE: duplicate_dim
' TYPE: prerun
' VBA: コンパイルエラー: 同じ適用範囲内で宣言が重複しています
' RUNNER: /duplicate/i
Sub Case_duplicate_dim()
    Dim v
    Dim v
End Sub

' ----------------------------------------------------------------
' 備考:
'   MySub(42) は VBE 上では自動フォーマットで `MySub (42)` に変換され、
'   コンパイルエラーにはならない（42 を括弧式として評価し Sub に渡す VBA 仕様）。
'   VBARunner での扱いは vba_compile_error.test.ts の skip コメントを参照。
' ----------------------------------------------------------------
