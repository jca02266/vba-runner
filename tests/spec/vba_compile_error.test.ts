/**
 * VBA コンパイルエラー自動テスト
 * このファイルは CompileError.bas から自動生成されました。
 * 再生成: npx tsx test-libs/compile-error-generator.ts tests/vba/CompileError.bas --output <このファイルのパス>
 *
 * [parse]  Parser.parse() 時に例外が発生するケース（行番号も検証）
 * [prerun] プロシージャ呼び出し直前の静的チェックで例外が発生するケース
 */

import { assertCompileErrorPass1, assertCompileErrorPass2 } from '../../test-libs/test-runner';
import { Lexer } from '../../src/engine/lexer';
import { Parser } from '../../src/engine/parser';
import { Evaluator } from '../../src/engine/evaluator';

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

// [prerun] assign_from_sub
// VBA: コンパイルエラー: FunctionまたはVariableが必要です
// VBA error line (within Sub body): 2
{
    try {
        assertCompileErrorPass2(`
      Private Sub MySub()
      End Sub
      
      Private Function MyFuncHasArg(x)
      End Function
      
      Sub __test__()
        Dim v
        v = MySub
      End Sub
      __test__
    `, 10, /function or variable/i, 'assign_from_sub');
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
        assertCompileErrorPass2(`
      Private Sub MySub()
      End Sub
      
      Private Function MyFuncHasArg(x)
      End Function
      
      Sub __test__()
        Dim v
        v = MySub()
      End Sub
      __test__
    `, 10, /function or variable/i, 'assign_from_sub_with_parens');
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
        assertCompileErrorPass2(`
      Private Sub MySub()
      End Sub
      
      Private Function MyFuncHasArg(x)
      End Function
      
      Sub __test__()
        Dim v
        Dim v
      End Sub
      __test__
    `, 10, /duplicate/i, 'duplicate_dim');
        console.log('[PASS] duplicate_dim');
        __pass__++;
    } catch (e: any) {
        console.error('[FAIL] duplicate_dim:', e.message);
        __fail__++;
    }
}

// [prerun] goto_undefined_label
// VBA: コンパイルエラー: 行ラベルが定義されていません
// VBA error line (within Sub body): 1
{
    try {
        assertCompileErrorPass2(`
      Private Sub MySub()
      End Sub
      
      Private Function MyFuncHasArg(x)
      End Function
      
      Sub __test__()
        GoTo NoSuchLabel
      End Sub
      __test__
    `, 9, /not defined.*label|label.*not defined/i, 'goto_undefined_label');
        console.log('[PASS] goto_undefined_label');
        __pass__++;
    } catch (e: any) {
        console.error('[FAIL] goto_undefined_label:', e.message);
        __fail__++;
    }
}

// [prerun] undefined_sub_call
// VBA: コンパイルエラー: SubまたはFunctionが定義されていません
// VBA error line (within Sub body): 1
{
    try {
        assertCompileErrorPass2(`
      Private Sub MySub()
      End Sub
      
      Private Function MyFuncHasArg(x)
      End Function
      
      Sub __test__()
        UnknownProc
      End Sub
      __test__
    `, 9, /sub or function not defined/i, 'undefined_sub_call');
        console.log('[PASS] undefined_sub_call');
        __pass__++;
    } catch (e: any) {
        console.error('[FAIL] undefined_sub_call:', e.message);
        __fail__++;
    }
}

// [prerun] qualified_undeclared_obj
// VBA: コンパイルエラー: 変数が定義されていません（Option Explicit で unknownModule が未宣言）
// OE 違反は callProcedure 時に throw される設計（Pass2 はマップ構築のみ）
// VBA error line (within Sub body): 3
{
    try {
        const src = `
      Private Sub MySub()
      End Sub

      Private Function MyFuncHasArg(x)
      End Function

      Option Explicit
      Sub Case_qualified_undeclared_obj()
          UnknownModule.UnknownProc
      End Sub
    `;
        const ast = new Parser(new Lexer(src).tokenize()).parse();
        const ev = new Evaluator(console.log);
        ev.evaluateModule(ast);
        ev.resolveIdentifiers([{ ast, moduleName: '' }]);
        let threw = false;
        let msg = '';
        try {
            ev.callProcedure('Case_qualified_undeclared_obj', []);
        } catch (e: any) {
            threw = true;
            msg = e?.message ?? String(e);
        }
        if (!threw) throw new Error('[FAIL] qualified_undeclared_obj: Expected error but none was thrown');
        if (!/variable not declared|not declared/i.test(msg)) throw new Error(`[FAIL] qualified_undeclared_obj: Message mismatch: "${msg}"`);
        if (!/\bline 10\b/.test(msg)) throw new Error(`[FAIL] qualified_undeclared_obj: Line mismatch, got: "${msg}"`);
        console.log('[PASS] qualified_undeclared_obj');
        __pass__++;
    } catch (e: any) {
        console.error('[FAIL] qualified_undeclared_obj:', e.message);
        __fail__++;
    }
}

// [prerun] sub_call_arg_count_mismatch
// VBA: コンパイルエラー: 引数の数が一致していません。または不正なプロパティを指定しています。
// VBA error line (within Sub body): 1
{
    try {
        assertCompileErrorPass2(`
      Private Sub MySub()
      End Sub
      
      Private Function MyFuncHasArg(x)
      End Function
      
      Sub __test__()
        MySub (42)
      End Sub
      __test__
    `, 9, /wrong number of arguments/i, 'sub_call_arg_count_mismatch');
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
        assertCompileErrorPass2(`
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
console.log(`\n=== Summary: ${__pass__} passed, ${__fail__} failed ===`);
if (__fail__ > 0) process.exit(1);
