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
    v = MySub ' @error
End Sub

' CASE: assign_from_sub_with_parens
' TYPE: prerun
' VBA: コンパイルエラー: FunctionまたはVariableが必要です
' RUNNER: /function or variable/i
Sub Case_assign_from_sub_with_parens()
    Dim v
    v = MySub() ' @error
End Sub

' CASE: duplicate_dim
' TYPE: prerun
' VBA: コンパイルエラー: 同じ適用範囲内で宣言が重複しています
' RUNNER: /duplicate/i
Sub Case_duplicate_dim()
    Dim v
    Dim v ' @error
End Sub

' CASE: goto_undefined_label
' TYPE: prerun
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
' VBARunner の TYPE 分類の違い:
'   同一モジュール内のプロシージャ名重複 → VBE では prerun だが VBARunner では parse 時に検出
' ----------------------------------------------------------------
' CASE: sub_call_arg_count_mismatch
' TYPE: prerun
' VBA: コンパイルエラー: 引数の数が一致していません。または不正なプロパティを指定しています。
' RUNNER: /wrong number of arguments/i
' NOTE: MySub(42) は VBE 上では自動フォーマットで `MySub (42)` に変換され
'   コンパイルエラーにはならない（42 を括弧式として評価し Sub に渡す VBA 仕様）。
Sub Case_sub_call_arg_count_mismatch()
    MySub (42)  ' @error
End Sub

' CASE: duplicate_sub_name
' TYPE: parse
' VBA: コンパイルエラー: 名前が適切ではありません
' RUNNER: /duplicate.*procedure|duplicate.*name/i
' NOTE: 同一モジュール内のプロシージャ名重複
'@case-begin
Sub duplicate_sub_name()

End Sub
''Sub duplicate_sub_name()  ' @error
''
''End Sub
'@case-end

