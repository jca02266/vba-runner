/**
 * VBA コンパイルエラー自動テスト
 * このファイルは CompileError.bas から自動生成されました。
 * 再生成: npx tsx test-libs/compile-error-generator.ts tests/vba/CompileError.bas --output <このファイルのパス>
 *
 * [parse]  Parser.parse() 時に例外が発生するケース（行番号も検証）
 * [prerun] resolveIdentifiers で例外が発生するケース（実行なし）
 * [preproc] precheckProc（OE チェック）で例外が発生するケース
 * [exec]  precheckProc 後の実行中に例外が発生するケース
 */

import { assertCompileErrorPass1, assertCompileErrorPrerun, assertCompileErrorPreproc, assertCompileErrorExec, captureCompileErrorMessage } from '../../test-libs/test-runner';

let __pass__ = 0, __fail__ = 0;

// [parse] sub_call_with_empty_parens1
// VBA: コンパイルエラー: 構文エラー
// VBA error line (within Sub body): 1
{
    try {
        assertCompileErrorPass1(`MySub()`, 1, /syntax error|parse error/i, 'sub_call_with_empty_parens1');
        console.log('[PASS] sub_call_with_empty_parens1');
        __pass__++;
    } catch (e: any) {
        console.error('[FAIL] sub_call_with_empty_parens1:', e.message);
        __fail__++;
    }
}

// [parse] sub_call_with_empty_parens2
// VBA: コンパイルエラー: 構文エラー
// VBA error line (within Sub body): 1
{
    try {
        assertCompileErrorPass1(`MyFunction()`, 1, /syntax error|parse error/i, 'sub_call_with_empty_parens2');
        console.log('[PASS] sub_call_with_empty_parens2');
        __pass__++;
    } catch (e: any) {
        console.error('[FAIL] sub_call_with_empty_parens2:', e.message);
        __fail__++;
    }
}

// [parse] sub_call_with_empty_parens3
// VBA: コンパイルエラー: 構文エラー
// VBA error line (within Sub body): 1
{
    try {
        assertCompileErrorPass1(`MySub ()`, 1, /syntax error|parse error/i, 'sub_call_with_empty_parens3');
        console.log('[PASS] sub_call_with_empty_parens3');
        __pass__++;
    } catch (e: any) {
        console.error('[FAIL] sub_call_with_empty_parens3:', e.message);
        __fail__++;
    }
}

// [parse] sub_call_with_empty_parens4
// VBA: コンパイルエラー: 構文エラー
// VBA error line (within Sub body): 1
{
    try {
        assertCompileErrorPass1(`UnKnown()`, 1, /syntax error|parse error/i, 'sub_call_with_empty_parens4');
        console.log('[PASS] sub_call_with_empty_parens4');
        __pass__++;
    } catch (e: any) {
        console.error('[FAIL] sub_call_with_empty_parens4:', e.message);
        __fail__++;
    }
}

// [parse] sub_call_arg_without_call_keyword
// VBA: コンパイルエラー: 構文エラー
// VBA error line (within Sub body): 1
{
    try {
        assertCompileErrorPass1(`Call MySub 42`, 1, /syntax error|parse error/i, 'sub_call_arg_without_call_keyword');
        console.log('[PASS] sub_call_arg_without_call_keyword');
        __pass__++;
    } catch (e: any) {
        console.error('[FAIL] sub_call_arg_without_call_keyword:', e.message);
        __fail__++;
    }
}

// [parse] assign_func_arg_no_parens
// VBA: コンパイルエラー: 構文エラー
// VBA error line (within Sub body): 2
{
    try {
        assertCompileErrorPass1(`Dim v
v = MyFuncHasArg arg`, 2, /syntax error|parse error/i, 'assign_func_arg_no_parens');
        console.log('[PASS] assign_func_arg_no_parens');
        __pass__++;
    } catch (e: any) {
        console.error('[FAIL] assign_func_arg_no_parens:', e.message);
        __fail__++;
    }
}

// [parse] label_then_sub_call_with_empty_parens1
// VBA: コンパイルエラー: 構文エラー
// VBA error line (within Sub body): 1
{
    try {
        assertCompileErrorPass1(`MySub: MySub()`, 1, /syntax error|parse error/i, 'label_then_sub_call_with_empty_parens1');
        console.log('[PASS] label_then_sub_call_with_empty_parens1');
        __pass__++;
    } catch (e: any) {
        console.error('[FAIL] label_then_sub_call_with_empty_parens1:', e.message);
        __fail__++;
    }
}

// [parse] label_then_sub_call_with_empty_parens2
// VBA: コンパイルエラー: 構文エラー
// VBA error line (within Sub body): 1
{
    try {
        assertCompileErrorPass1(`MySub : MySub()`, 1, /syntax error|parse error/i, 'label_then_sub_call_with_empty_parens2');
        console.log('[PASS] label_then_sub_call_with_empty_parens2');
        __pass__++;
    } catch (e: any) {
        console.error('[FAIL] label_then_sub_call_with_empty_parens2:', e.message);
        __fail__++;
    }
}

// [parse] reserved_word_as_function_name
// VBA: コンパイルエラー: 修正候補: 識別子
// VBA error line (within Sub body): 1
{
    try {
        assertCompileErrorPass1(`Function Open() As Long
    Open = 1
End Function`, 1, /reserved word/i, 'reserved_word_as_function_name');
        console.log('[PASS] reserved_word_as_function_name');
        __pass__++;
    } catch (e: any) {
        console.error('[FAIL] reserved_word_as_function_name:', e.message);
        __fail__++;
    }
}

// [parse] reserved_word_as_sub_name
// VBA: コンパイルエラー: 修正候補: 識別子
// VBA error line (within Sub body): 1
{
    try {
        assertCompileErrorPass1(`Sub Close()
End Sub`, 1, /reserved word/i, 'reserved_word_as_sub_name');
        console.log('[PASS] reserved_word_as_sub_name');
        __pass__++;
    } catch (e: any) {
        console.error('[FAIL] reserved_word_as_sub_name:', e.message);
        __fail__++;
    }
}

// [parse] reserved_word_as_function_name_print
// VBA: コンパイルエラー: 修正候補: 識別子
// VBA error line (within Sub body): 1
{
    try {
        assertCompileErrorPass1(`Function Print() As Long
    Print = 99
End Function`, 1, /reserved word/i, 'reserved_word_as_function_name_print');
        console.log('[PASS] reserved_word_as_function_name_print');
        __pass__++;
    } catch (e: any) {
        console.error('[FAIL] reserved_word_as_function_name_print:', e.message);
        __fail__++;
    }
}

// [preproc] assign_from_sub
// VBA: コンパイルエラー: FunctionまたはVariableが必要です
// VBA error line (within Sub body): 2
{
    try {
        assertCompileErrorPreproc(`
      Private Sub MySub()
      End Sub
      
      Private Function MyFuncHasArg(x)
      End Function
      
      Sub __test__()
        Dim v
        v = MySub
      End Sub
    `, '__test__', 10, /function or variable/i, 'assign_from_sub');
        console.log('[PASS] assign_from_sub');
        __pass__++;
    } catch (e: any) {
        console.error('[FAIL] assign_from_sub:', e.message);
        __fail__++;
    }
}

// [preproc] assign_from_sub_with_parens
// VBA: コンパイルエラー: FunctionまたはVariableが必要です
// VBA error line (within Sub body): 2
{
    try {
        assertCompileErrorPreproc(`
      Private Sub MySub()
      End Sub
      
      Private Function MyFuncHasArg(x)
      End Function
      
      Sub __test__()
        Dim v
        v = MySub()
      End Sub
    `, '__test__', 10, /function or variable/i, 'assign_from_sub_with_parens');
        console.log('[PASS] assign_from_sub_with_parens');
        __pass__++;
    } catch (e: any) {
        console.error('[FAIL] assign_from_sub_with_parens:', e.message);
        __fail__++;
    }
}

// [preproc] duplicate_dim
// VBA: コンパイルエラー: 同じ適用範囲内で宣言が重複しています
// VBA error line (within Sub body): 2
{
    try {
        assertCompileErrorPreproc(`
      Private Sub MySub()
      End Sub
      
      Private Function MyFuncHasArg(x)
      End Function
      
      Sub __test__()
        Dim v
        Dim v
      End Sub
    `, '__test__', 10, /duplicate/i, 'duplicate_dim');
        console.log('[PASS] duplicate_dim');
        __pass__++;
    } catch (e: any) {
        console.error('[FAIL] duplicate_dim:', e.message);
        __fail__++;
    }
}

// [preproc] goto_undefined_label
// VBA: コンパイルエラー: 行ラベルが定義されていません
// VBA error line (within Sub body): 1
{
    try {
        assertCompileErrorPreproc(`
      Private Sub MySub()
      End Sub
      
      Private Function MyFuncHasArg(x)
      End Function
      
      Sub __test__()
        GoTo NoSuchLabel
      End Sub
    `, '__test__', 9, /not defined.*label|label.*not defined/i, 'goto_undefined_label');
        console.log('[PASS] goto_undefined_label');
        __pass__++;
    } catch (e: any) {
        console.error('[FAIL] goto_undefined_label:', e.message);
        __fail__++;
    }
}

// [preproc] undefined_sub_call
// VBA: コンパイルエラー: SubまたはFunctionが定義されていません
// VBA error line (within Sub body): 1
{
    try {
        assertCompileErrorPreproc(`
      Private Sub MySub()
      End Sub
      
      Private Function MyFuncHasArg(x)
      End Function
      
      Sub __test__()
        UnknownProc
      End Sub
    `, '__test__', 9, /sub or function not defined/i, 'undefined_sub_call');
        console.log('[PASS] undefined_sub_call');
        __pass__++;
    } catch (e: any) {
        console.error('[FAIL] undefined_sub_call:', e.message);
        __fail__++;
    }
}

// [preproc] qualified_undeclared_obj
// VBA: コンパイルエラー: 変数が定義されていません（Option Explicit で unknownModule が未宣言）
// VBA error line (within Sub body): 3
{
    try {
        assertCompileErrorPreproc(`
      Private Sub MySub()
      End Sub
      
      Private Function MyFuncHasArg(x)
      End Function
      
      Option Explicit
      Sub Case_qualified_undeclared_obj()
          UnknownModule.UnknownProc
      End Sub
    `, 'Case_qualified_undeclared_obj', 10, /variable not declared|not declared/i, 'qualified_undeclared_obj');
        console.log('[PASS] qualified_undeclared_obj');
        __pass__++;
    } catch (e: any) {
        console.error('[FAIL] qualified_undeclared_obj:', e.message);
        __fail__++;
    }
}

// [preproc] undefined_sub_call_no_oe
// VBA: コンパイルエラー: SubまたはFunctionが定義されていません
// VBA error line (within Sub body): 2
{
    try {
        assertCompileErrorPreproc(`
      Private Sub MySub()
      End Sub
      
      Private Function MyFuncHasArg(x)
      End Function
      
      Sub Case_undefined_sub_call_no_oe()
          UnknownProc
      End Sub
    `, 'Case_undefined_sub_call_no_oe', 9, /sub or function not defined/i, 'undefined_sub_call_no_oe');
        console.log('[PASS] undefined_sub_call_no_oe');
        __pass__++;
    } catch (e: any) {
        console.error('[FAIL] undefined_sub_call_no_oe:', e.message);
        __fail__++;
    }
}

// [preproc] undefined_sub_call_with_oe
// VBA: コンパイルエラー: SubまたはFunctionが定義されていません
// VBA error line (within Sub body): 3
{
    try {
        assertCompileErrorPreproc(`
      Private Sub MySub()
      End Sub
      
      Private Function MyFuncHasArg(x)
      End Function
      
      Option Explicit
      Sub Case_undefined_sub_call_with_oe()
          UnknownProc
      End Sub
    `, 'Case_undefined_sub_call_with_oe', 10, /sub or function not defined/i, 'undefined_sub_call_with_oe');
        console.log('[PASS] undefined_sub_call_with_oe');
        __pass__++;
    } catch (e: any) {
        console.error('[FAIL] undefined_sub_call_with_oe:', e.message);
        __fail__++;
    }
}

// [preproc] sub_call_arg_count_mismatch
// VBA: コンパイルエラー: 引数の数が一致していません。または不正なプロパティを指定しています。
// VBA error line (within Sub body): 1
{
    try {
        assertCompileErrorPreproc(`
      Private Sub MySub()
      End Sub
      
      Private Function MyFuncHasArg(x)
      End Function
      
      Sub __test__()
        MySub (42)
      End Sub
    `, '__test__', 9, /wrong number of arguments/i, 'sub_call_arg_count_mismatch');
        console.log('[PASS] sub_call_arg_count_mismatch');
        __pass__++;
    } catch (e: any) {
        console.error('[FAIL] sub_call_arg_count_mismatch:', e.message);
        __fail__++;
    }
}

// [prerun] duplicate_sub_name
// VBA: コンパイルエラー: 名前が適切ではありません duplicate_sub_name
// VBA error line (within Sub body): 4
{
    try {
        assertCompileErrorPrerun(`
      Private Sub MySub()
      End Sub
      
      Private Function MyFuncHasArg(x)
      End Function
      
      Sub duplicate_sub_name()
      
      End Sub
      Sub duplicate_sub_name()
      
      End Sub
    `, 11, /duplicate.*procedure|duplicate.*name/i, 'duplicate_sub_name');
        console.log('[PASS] duplicate_sub_name');
        __pass__++;
    } catch (e: any) {
        console.error('[FAIL] duplicate_sub_name:', e.message);
        __fail__++;
    }
}

// [prerun] module_level_dim_after_procedure
// VBA: コンパイル エラー: End Sub、End Function または End Property 以降には、コメントのみが記述できます。
// VBA error line (within Sub body): 3
{
    try {
        assertCompileErrorPrerun(`
      Private Sub MySub()
      End Sub
      
      Private Function MyFuncHasArg(x)
      End Function
      
      Sub ModuleLevelDimAfterProcedure()
      End Sub
      Dim moduleLevelVar As Integer
    `, 10, /only comments may appear after end sub/i, 'module_level_dim_after_procedure', { allowTopLevelStatements: false });
        console.log('[PASS] module_level_dim_after_procedure');
        __pass__++;
    } catch (e: any) {
        console.error('[FAIL] module_level_dim_after_procedure:', e.message);
        __fail__++;
    }
}

// [prerun] module_level_const_after_procedure
// VBA: コンパイル エラー: End Sub、End Function または End Property 以降には、コメントのみが記述できます。
// VBA error line (within Sub body): 3
{
    try {
        assertCompileErrorPrerun(`
      Private Sub MySub()
      End Sub
      
      Private Function MyFuncHasArg(x)
      End Function
      
      Sub ModuleLevelConstAfterProcedure()
      End Sub
      Const ModuleLevelConst As Integer = 1
    `, 10, /only comments may appear after end sub/i, 'module_level_const_after_procedure', { allowTopLevelStatements: false });
        console.log('[PASS] module_level_const_after_procedure');
        __pass__++;
    } catch (e: any) {
        console.error('[FAIL] module_level_const_after_procedure:', e.message);
        __fail__++;
    }
}

// [prerun] module_level_public_after_procedure
// VBA: コンパイル エラー: End Sub、End Function または End Property 以降には、コメントのみが記述できます。
// VBA error line (within Sub body): 3
{
    try {
        assertCompileErrorPrerun(`
      Private Sub MySub()
      End Sub
      
      Private Function MyFuncHasArg(x)
      End Function
      
      Sub ModuleLevelPublicAfterProcedure()
      End Sub
      Public ModuleLevelPublicVar As Integer
    `, 10, /only comments may appear after end sub/i, 'module_level_public_after_procedure', { allowTopLevelStatements: false });
        console.log('[PASS] module_level_public_after_procedure');
        __pass__++;
    } catch (e: any) {
        console.error('[FAIL] module_level_public_after_procedure:', e.message);
        __fail__++;
    }
}

// [prerun] module_level_type_after_procedure
// VBA: コンパイル エラー: End Sub、End Function または End Property 以降には、コメントのみが記述できます。
// VBA error line (within Sub body): 3
{
    try {
        assertCompileErrorPrerun(`
      Private Sub MySub()
      End Sub
      
      Private Function MyFuncHasArg(x)
      End Function
      
      Sub ModuleLevelTypeAfterProcedure()
      End Sub
      Type ModuleLevelType
          Field As Integer
      End Type
    `, 10, /only comments may appear after end sub/i, 'module_level_type_after_procedure', { allowTopLevelStatements: false });
        console.log('[PASS] module_level_type_after_procedure');
        __pass__++;
    } catch (e: any) {
        console.error('[FAIL] module_level_type_after_procedure:', e.message);
        __fail__++;
    }
}

// [prerun] module_level_enum_after_procedure
// VBA: コンパイル エラー: End Sub、End Function または End Property 以降には、コメントのみが記述できます。
// VBA error line (within Sub body): 3
{
    try {
        assertCompileErrorPrerun(`
      Private Sub MySub()
      End Sub
      
      Private Function MyFuncHasArg(x)
      End Function
      
      Sub ModuleLevelEnumAfterProcedure()
      End Sub
      Enum ModuleLevelEnum
          ModuleLevelEnumValue
      End Enum
    `, 10, /only comments may appear after end sub/i, 'module_level_enum_after_procedure', { allowTopLevelStatements: false });
        console.log('[PASS] module_level_enum_after_procedure');
        __pass__++;
    } catch (e: any) {
        console.error('[FAIL] module_level_enum_after_procedure:', e.message);
        __fail__++;
    }
}

// [prerun] module_level_toplevel_stmt_after_procedure_strict
// VBA: コンパイル エラー: End Sub、End Function または End Property 以降には、コメントのみが記述できます。
// VBA error line (within Sub body): 3
{
    try {
        assertCompileErrorPrerun(`
      Private Sub MySub()
      End Sub
      
      Private Function MyFuncHasArg(x)
      End Function
      
      Sub ModuleLevelToplevelStmtAfterProcedureStrict()
      End Sub
      For moduleLevelIdx = 0 To 10
      Next moduleLevelIdx
    `, 10, /only comments may appear after end sub/i, 'module_level_toplevel_stmt_after_procedure_strict', { allowTopLevelStatements: false });
        console.log('[PASS] module_level_toplevel_stmt_after_procedure_strict');
        __pass__++;
    } catch (e: any) {
        console.error('[FAIL] module_level_toplevel_stmt_after_procedure_strict:', e.message);
        __fail__++;
    }
}
console.log(`\n=== Summary: ${__pass__} passed, ${__fail__} failed ===`);
if (__fail__ > 0) process.exit(1);
