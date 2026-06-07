/**
 * VBA コンパイルエラー自動テスト
 * このファイルは CompileError.bas から自動生成されました。
 * 再生成: npx tsx test-libs/compile-error-generator.ts tests/vba/CompileError.bas --output <このファイルのパス>
 *
 * [parse]  Parser.parse() 時に例外が発生するケース
 * [prerun] プロシージャ呼び出し直前の静的チェックで例外が発生するケース
 */

import { Lexer } from '../../src/engine/lexer';
import { Parser } from '../../src/engine/parser';
import { evalVBASingle, assert } from '../../test-libs/test-runner';

let __pass__ = 0, __fail__ = 0;

// [parse] sub_call_with_empty_parens
// VBA: コンパイルエラー: 構文エラー
{
    try {
        const src = `
    MySub()
`;
        assert.throwsMatch(() => {
            const tokens = new Lexer(src).tokenize();
            new Parser(tokens).parse();
        }, /syntax error|parse error/i, 'sub_call_with_empty_parens');
        console.log('[PASS] sub_call_with_empty_parens');
        __pass__++;
    } catch (e: any) {
        console.error('[FAIL] sub_call_with_empty_parens:', e.message);
        __fail__++;
    }
}

// [parse] sub_call_arg_without_call_keyword
// VBA: コンパイルエラー: 構文エラー
{
    try {
        const src = `
    Call MySub 42
`;
        assert.throwsMatch(() => {
            const tokens = new Lexer(src).tokenize();
            new Parser(tokens).parse();
        }, /syntax error|parse error/i, 'sub_call_arg_without_call_keyword');
        console.log('[PASS] sub_call_arg_without_call_keyword');
        __pass__++;
    } catch (e: any) {
        console.error('[FAIL] sub_call_arg_without_call_keyword:', e.message);
        __fail__++;
    }
}

// [parse] assign_func_arg_no_parens
// VBA: コンパイルエラー: 構文エラー
{
    try {
        const src = `
    Dim v
    v = MyFuncHasArg arg
`;
        assert.throwsMatch(() => {
            const tokens = new Lexer(src).tokenize();
            new Parser(tokens).parse();
        }, /syntax error|parse error/i, 'assign_func_arg_no_parens');
        console.log('[PASS] assign_func_arg_no_parens');
        __pass__++;
    } catch (e: any) {
        console.error('[FAIL] assign_func_arg_no_parens:', e.message);
        __fail__++;
    }
}

// [prerun] assign_from_sub
// VBA: コンパイルエラー: FunctionまたはVariableが必要です
{
    try {
        assert.throwsMatch(() => evalVBASingle(`
      Private Sub MySub()
      End Sub
      
      Private Function MyFuncHasArg(x)
      End Function
      
      Sub __test__()
        Dim v
        v = MySub
      End Sub
      __test__
    `), /function or variable/i, 'assign_from_sub');
        console.log('[PASS] assign_from_sub');
        __pass__++;
    } catch (e: any) {
        console.error('[FAIL] assign_from_sub:', e.message);
        __fail__++;
    }
}

// [prerun] assign_from_sub_with_parens
// VBA: コンパイルエラー: FunctionまたはVariableが必要です
{
    try {
        assert.throwsMatch(() => evalVBASingle(`
      Private Sub MySub()
      End Sub
      
      Private Function MyFuncHasArg(x)
      End Function
      
      Sub __test__()
        Dim v
        v = MySub()
      End Sub
      __test__
    `), /function or variable/i, 'assign_from_sub_with_parens');
        console.log('[PASS] assign_from_sub_with_parens');
        __pass__++;
    } catch (e: any) {
        console.error('[FAIL] assign_from_sub_with_parens:', e.message);
        __fail__++;
    }
}

// [prerun] duplicate_dim
// VBA: コンパイルエラー: 同じ適用範囲内で宣言が重複しています
{
    try {
        assert.throwsMatch(() => evalVBASingle(`
      Private Sub MySub()
      End Sub
      
      Private Function MyFuncHasArg(x)
      End Function
      
      Sub __test__()
        Dim v
        Dim v
      End Sub
      __test__
    `), /duplicate/i, 'duplicate_dim');
        console.log('[PASS] duplicate_dim');
        __pass__++;
    } catch (e: any) {
        console.error('[FAIL] duplicate_dim:', e.message);
        __fail__++;
    }
}
console.log(`\n=== Summary: ${__pass__} passed, ${__fail__} failed ===`);
if (__fail__ > 0) process.exit(1);
