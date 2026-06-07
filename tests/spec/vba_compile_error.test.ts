/**
 * VBA コンパイルエラー自動テスト
 * このファイルは CompileError.bas から自動生成されました。
 * 再生成: npx tsx test-libs/compile-error-generator.ts tests/vba/CompileError.bas --output <このファイルのパス>
 *
 * [parse]  Parser.parse() 時に例外が発生するケース（行番号も検証）
 * [prerun] プロシージャ呼び出し直前の静的チェックで例外が発生するケース
 */

import { Lexer } from '../../src/engine/lexer';
import { Parser } from '../../src/engine/parser';
import { evalVBASingle, assertCompileError } from '../../test-libs/test-runner';

let __pass__ = 0, __fail__ = 0;

// [parse] sub_call_with_empty_parens1
// VBA: コンパイルエラー: 構文エラー
// VBA error line (within Sub body): 1
{
    try {
        const src = `MySub()`;
        assertCompileError(
            () => { new Parser(new Lexer(src).tokenize()).parse(); },
            1, /syntax error|parse error/i, 'sub_call_with_empty_parens1'
        );
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
        const src = `MyFunction()`;
        assertCompileError(
            () => { new Parser(new Lexer(src).tokenize()).parse(); },
            1, /syntax error|parse error/i, 'sub_call_with_empty_parens2'
        );
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
        const src = `MySub ()`;
        assertCompileError(
            () => { new Parser(new Lexer(src).tokenize()).parse(); },
            1, /syntax error|parse error/i, 'sub_call_with_empty_parens3'
        );
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
        const src = `UnKnown()`;
        assertCompileError(
            () => { new Parser(new Lexer(src).tokenize()).parse(); },
            1, /syntax error|parse error/i, 'sub_call_with_empty_parens4'
        );
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
        const src = `Call MySub 42`;
        assertCompileError(
            () => { new Parser(new Lexer(src).tokenize()).parse(); },
            1, /syntax error|parse error/i, 'sub_call_arg_without_call_keyword'
        );
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
        const src = `Dim v
v = MyFuncHasArg arg`;
        assertCompileError(
            () => { new Parser(new Lexer(src).tokenize()).parse(); },
            2, /syntax error|parse error/i, 'assign_func_arg_no_parens'
        );
        console.log('[PASS] assign_func_arg_no_parens');
        __pass__++;
    } catch (e: any) {
        console.error('[FAIL] assign_func_arg_no_parens:', e.message);
        __fail__++;
    }
}

// [prerun] assign_from_sub
// VBA: コンパイルエラー: FunctionまたはVariableが必要です
// VBA error line (within Sub body): 2
{
    try {
        assertCompileError(() => evalVBASingle(`
      Private Sub MySub()
      End Sub
      
      Private Function MyFuncHasArg(x)
      End Function
      
      Sub __test__()
        Dim v
        v = MySub
      End Sub
      __test__
    `), 10, /function or variable/i, 'assign_from_sub');
        console.log('[PASS] assign_from_sub');
        __pass__++;
    } catch (e: any) {
        console.error('[FAIL] assign_from_sub:', e.message);
        __fail__++;
    }
}

// [prerun] assign_from_sub_with_parens
// VBA: コンパイルエラー: FunctionまたはVariableが必要です
// VBA error line (within Sub body): 2
{
    try {
        assertCompileError(() => evalVBASingle(`
      Private Sub MySub()
      End Sub
      
      Private Function MyFuncHasArg(x)
      End Function
      
      Sub __test__()
        Dim v
        v = MySub()
      End Sub
      __test__
    `), 10, /function or variable/i, 'assign_from_sub_with_parens');
        console.log('[PASS] assign_from_sub_with_parens');
        __pass__++;
    } catch (e: any) {
        console.error('[FAIL] assign_from_sub_with_parens:', e.message);
        __fail__++;
    }
}

// [prerun] duplicate_dim
// VBA: コンパイルエラー: 同じ適用範囲内で宣言が重複しています
// VBA error line (within Sub body): 2
{
    try {
        assertCompileError(() => evalVBASingle(`
      Private Sub MySub()
      End Sub
      
      Private Function MyFuncHasArg(x)
      End Function
      
      Sub __test__()
        Dim v
        Dim v
      End Sub
      __test__
    `), 10, /duplicate/i, 'duplicate_dim');
        console.log('[PASS] duplicate_dim');
        __pass__++;
    } catch (e: any) {
        console.error('[FAIL] duplicate_dim:', e.message);
        __fail__++;
    }
}
console.log(`\n=== Summary: ${__pass__} passed, ${__fail__} failed ===`);
if (__fail__ > 0) process.exit(1);
