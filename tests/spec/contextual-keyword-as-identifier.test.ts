/**
 * §3.3.5.2: IDENTIFIER として使える contextual keyword のテスト
 *
 * VBA 仕様上 reserved-identifier に属さないキーワードは IDENTIFIER として使える。
 * Lexer が専用トークン型を持つ場合でも、宣言名・変数名・定数名・プロシージャ名に
 * 使用できることを確認する。
 */
import { evalVBASingle, evalVBAModules, assert, assertCompileErrorPass1 } from '../../test-libs/test-runner';

function run(code: string, name = 'T'): any {
    return evalVBASingle(code).callProcedure(name, []);
}

// --- 1. Dim 宣言の変数名 ---
{
    // CONTEXTUAL_KW（既存）
    assert.strictEqual(run(`Function T() As Long: Dim Output As Long: Output = 1: T = Output: End Function`), 1, 'Dim Output');
    assert.strictEqual(run(`Function T() As Long: Dim Step As Long: Step = 2: T = Step: End Function`), 2, 'Dim Step');
    assert.strictEqual(run(`Function T() As Long: Dim Binary As Long: Binary = 3: T = Binary: End Function`), 3, 'Dim Binary');
    assert.strictEqual(run(`Function T() As Long: Dim Access As Long: Access = 4: T = Access: End Function`), 4, 'Dim Access');
    assert.strictEqual(run(`Function T() As String: Dim Text As String: Text = "hi": T = Text: End Function`), "hi", 'Dim Text');
    // CONTEXTUAL_KW_STRUCTURAL（追加分）
    assert.strictEqual(run(`Function T() As Long: Dim Error As Long: Error = 5: T = Error: End Function`), 5, 'Dim Error');
    assert.strictEqual(run(`Function T() As Long: Dim Class As Long: Class = 6: T = Class: End Function`), 6, 'Dim Class');
    assert.strictEqual(run(`Function T() As Long: Dim Property As Long: Property = 7: T = Property: End Function`), 7, 'Dim Property');
    assert.strictEqual(run(`Function T() As Long: Dim Collection As Long: Collection = 8: T = Collection: End Function`), 8, 'Dim Collection');
    console.log('[PASS] Dim 宣言の変数名');
}

// --- 2. Const 宣言の定数名 ---
{
    // CONTEXTUAL_KW（既存）
    assert.strictEqual(run(`Const Step = 10\nFunction T() As Long: T = Step: End Function`), 10, 'Const Step');
    assert.strictEqual(run(`Const Binary = 11\nFunction T() As Long: T = Binary: End Function`), 11, 'Const Binary');
    assert.strictEqual(run(`Const Output = 12\nFunction T() As Long: T = Output: End Function`), 12, 'Const Output');
    // CONTEXTUAL_KW_STRUCTURAL（追加分）
    assert.strictEqual(run(`Const Error = 13\nFunction T() As Long: T = Error: End Function`), 13, 'Const Error');
    assert.strictEqual(run(`Const Class = 14\nFunction T() As Long: T = Class: End Function`), 14, 'Const Class');
    assert.strictEqual(run(`Const Property = 15\nFunction T() As Long: T = Property: End Function`), 15, 'Const Property');
    assert.strictEqual(run(`Const Collection = 16\nFunction T() As Long: T = Collection: End Function`), 16, 'Const Collection');
    console.log('[PASS] Const 宣言の定数名');
}

// --- 3. Function/Sub 名 ---
{
    // CONTEXTUAL_KW（既存）
    assert.strictEqual(run(`Function Output() As Long: Output = 99: End Function`, 'Output'), 99, 'Function Output');
    assert.strictEqual(run(`Function Binary() As Long: Binary = 100: End Function`, 'Binary'), 100, 'Function Binary');
    // CONTEXTUAL_KW_STRUCTURAL（追加分）
    assert.strictEqual(run(`Function Error() As Long: Error = 101: End Function`, 'Error'), 101, 'Function Error');
    assert.strictEqual(run(`Function Property() As Long: Property = 102: End Function`, 'Property'), 102, 'Function Property');
    assert.strictEqual(run(`Function Class() As Long: Class = 103: End Function`, 'Class'), 103, 'Function Class');
    assert.strictEqual(run(`Function Collection() As Long: Collection = 104: End Function`, 'Collection'), 104, 'Function Collection');
    console.log('[PASS] Function/Sub 名');
}

// --- 4. パラメーター名 ---
{
    const ev1 = evalVBASingle(`Function F(ByVal text As String): F = text: End Function`, { onPrint: () => {} });
    assert.strictEqual(ev1.callProcedure('F', ['hi']), 'hi', 'ByVal text As String');

    const ev2 = evalVBASingle(`Function F(ByVal binary As Long): F = binary: End Function`, { onPrint: () => {} });
    assert.strictEqual(ev2.callProcedure('F', [3]), 3, 'ByVal binary As Long');

    const ev3 = evalVBASingle(`Function F(ByVal compare As Long): F = compare: End Function`, { onPrint: () => {} });
    assert.strictEqual(ev3.callProcedure('F', [7]), 7, 'ByVal compare As Long');

    const ev4 = evalVBASingle(`Function F(ByVal output As Long): F = output: End Function`, { onPrint: () => {} });
    assert.strictEqual(ev4.callProcedure('F', [42]), 42, 'ByVal output As Long');

    console.log('[PASS] パラメーター名に contextual keyword を使える');
}

// --- 5. 同名の VBA 組み込み文との共存（代入 vs 文ディスパッチ）---
{
    // Class = value（Class宣言文と区別できる）
    assert.strictEqual(run(`Function T() As Long: Dim Class As Long: Class = 55: T = Class: End Function`), 55, 'Class 代入');
    // Error = value（Error文と区別できる）
    assert.strictEqual(run(`Function T() As Long: Dim Error As Long: Error = 56: T = Error: End Function`), 56, 'Error 代入');
    // Property = value（Property Get/Set/Let と区別できる）
    assert.strictEqual(run(`Function T() As Long: Dim Property As Long: Property = 57: T = Property: End Function`), 57, 'Property 代入');
    console.log('[PASS] 組み込み文との共存');
}

// --- 6. statement-keyword はモジュールレベルのプロシージャ名に使えない（§3.3.5.2）---
// Open / Close / Print / Input 等は reserved-identifier であるため
// Function/Sub 宣言名として使うとコンパイルエラーになる。
{
    assertCompileErrorPass1(
        `Function Open() As Boolean\n    Open = True\nEnd Function`,
        1, /reserved word/, 'Function Open() はコンパイルエラー');

    assertCompileErrorPass1(
        `Function Close() As Boolean\n    Close = True\nEnd Function`,
        1, /reserved word/, 'Function Close() はコンパイルエラー');

    assertCompileErrorPass1(
        `Function Print() As Long\n    Print = 99\nEnd Function`,
        1, /reserved word/, 'Function Print() はコンパイルエラー');

    assertCompileErrorPass1(
        `Function Input() As Long\n    Input = 55\nEnd Function`,
        1, /reserved word/, 'Function Input() はコンパイルエラー');

    assertCompileErrorPass1(
        `Function Write() As Long\n    Write = 1\nEnd Function`,
        1, /reserved word/, 'Function Write() はコンパイルエラー');

    assertCompileErrorPass1(
        `Function Seek() As Long\n    Seek = 1\nEnd Function`,
        1, /reserved word/, 'Function Seek() はコンパイルエラー');

    assertCompileErrorPass1(
        `Function Lock() As Long\n    Lock = 1\nEnd Function`,
        1, /reserved word/, 'Function Lock() はコンパイルエラー');

    assertCompileErrorPass1(
        `Function Unlock() As Long\n    Unlock = 1\nEnd Function`,
        1, /reserved word/, 'Function Unlock() はコンパイルエラー');

    console.log('[PASS] statement-keyword をプロシージャ名に使うとコンパイルエラー');
}

console.log('\n[PASS] contextual-keyword-as-identifier: 全テスト通過');
