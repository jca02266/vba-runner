Option Explicit
' ================================================================
' CompileError.bas - VBA コンパイルエラー仕様記述ファイル
'
' 【VBA コンパイルエラーの 2 種類】
'
' [parse] 構文エラー（実行前）:
'   VBE で該当行が赤くなり、実行ボタンを押す前からエラーがわかる。
'   エラーダイアログ: "コンパイルエラー: 構文エラー" など。
'   VBARunner 対応: Parser.parse() 時に例外が発生（パース時エラー）。
'
' [prerun] 実行直前コンパイルエラー:
'   実行ボタンを押すと VBE カーソルが該当箇所に移動し、実行直前に
'   コンパイルエラーダイアログが表示される。On Error では捕捉できない。
'   VBARunner 対応: プロシージャ呼び出し直前の静的チェックで例外が発生。
'
' 【自動テスト生成】
'   このファイルの CASE ブロックを解析して vba_compile_error.test.ts を生成する。
'   生成コマンド: npx tsx test-libs/compile-error-generator.ts tests/vba/CompileError.bas
'
' 【CASE ブロックの書式】
'   ' CASE: <snake_case_name>
'   ' TYPE: parse | prerun
'   ' VBA: <VBE が表示するエラーメッセージ>
'   ' RUNNER: <VBARunner が投げる例外メッセージの正規表現>
'   ' CODE:
'   '   <コード行>
'   '   <コード行>
'
' 【PREAMBLE ブロック】
'   各テストケースのコードを評価する前に注入するヘルパー定義。
'   CASE の CODE はこれらのプロシージャを参照できる。
' ================================================================

' PREAMBLE:
'   Private Sub MySub()
'   End Sub
'
'   Private Function MyFuncHasArg(x)
'   End Function

' ----------------------------------------------------------------
' [parse] 構文エラー（VBE が行を赤くし、実行前にわかる）
' ----------------------------------------------------------------

' CASE: sub_call_with_empty_parens
' TYPE: parse
' VBA: コンパイルエラー: 構文エラー
' RUNNER: /syntax error/i
' CODE:
'   MySub()

' CASE: sub_call_arg_without_call_keyword
' TYPE: parse
' VBA: コンパイルエラー: 構文エラー
' RUNNER: /syntax error/i
' CODE:
'   Call MySub 42

' CASE: assign_func_arg_no_parens
' TYPE: parse
' VBA: コンパイルエラー: 構文エラー
' RUNNER: /syntax error/i
' CODE:
'   Dim v
'   v = MyFuncHasArg arg

' ----------------------------------------------------------------
' [prerun] 実行直前コンパイルエラー（実行ボタンを押したときに検出）
' ----------------------------------------------------------------

' CASE: assign_from_sub
' TYPE: prerun
' VBA: コンパイルエラー: FunctionまたはVariableが必要です
' RUNNER: /function or variable/i
' CODE:
'   Dim v
'   v = MySub

' CASE: assign_from_sub_with_parens
' TYPE: prerun
' VBA: コンパイルエラー: FunctionまたはVariableが必要です
' RUNNER: /function or variable/i
' CODE:
'   Dim v
'   v = MySub()

' CASE: duplicate_dim
' TYPE: prerun
' VBA: コンパイルエラー: 同じ適用範囲内で宣言が重複しています
' RUNNER: /duplicate/i
' CODE:
'   Dim v
'   Dim v

' ----------------------------------------------------------------
' 備考:
'   MySub(42) は VBE 上では自動フォーマットで `MySub (42)` に変換され、
'   コンパイルエラーにはならない（42 を括弧式として評価し Sub に渡す VBA 仕様）。
'   VBARunner での扱いは vba_compile_error.test.ts の skip コメントを参照。
' ----------------------------------------------------------------
