/**
 * If 条件式の truthy 判定と Boolean 型強制の挙動確認テスト
 *
 * VBA の If 文は条件値を Boolean に暗黙変換して評価する。
 * 変換規則:
 *   - 数値: 0 → False、非 0 → True
 *   - VbaBoolean: そのまま
 *   - 文字列 "True"/"False": 対応する Boolean
 *   - 文字列（数値表現）: 数値化して評価
 *   - 文字列（それ以外）: Type mismatch (Error 13)
 *   - Empty (Null JS): False（未初期化 Variant = 0 = False）
 *   - Null: Invalid use of Null (Error 94)
 */
import { Lexer } from '../../src/engine/lexer';
import { Parser } from '../../src/engine/parser';
import { Evaluator, vbaTrue, vbaFalse, vbaNull } from '../../src/engine/evaluator';
import { assert } from '../../test-libs/test-runner';

function runFunc(code: string, name: string, args: any[] = []): any {
    const tokens = new Lexer(code).tokenize();
    const ast = new Parser(tokens).parse();
    const ev = new Evaluator(console.log);
    ev.evaluate(ast);
    return ev.callProcedure(name, args);
}

function expectError(fn: () => any, errorNumber: number, label: string) {
    let caught: any = null;
    try { fn(); } catch (e: any) { caught = e; }
    if (!caught) {
        console.error(`[FAIL] ${label} - エラーが発生しなかった`);
        throw new Error('Assertion Failed');
    }
    if (Number(caught.number) !== errorNumber) {
        console.error(`[FAIL] ${label} - 期待エラー ${errorNumber}、実際: ${JSON.stringify(caught)}`);
        throw new Error('Assertion Failed');
    }
}

const branch = (cond: string) => `
    Function Check() As String
        If ${cond} Then
            Check = "true"
        Else
            Check = "false"
        End If
    End Function
`;

console.log('--- Starting If-Condition Truthy Tests ---');

// =============================================================================
// 1. 数値の truthy 判定
// =============================================================================
{
    assert.strictEqual(runFunc(branch('0'), 'Check'),     'false', '0 → False');
    assert.strictEqual(runFunc(branch('1'), 'Check'),     'true',  '1 → True');
    assert.strictEqual(runFunc(branch('-1'), 'Check'),    'true',  '-1 → True（True の内部値）');
    assert.strictEqual(runFunc(branch('2'), 'Check'),     'true',  '2 → True（非 0）');
    assert.strictEqual(runFunc(branch('0.5'), 'Check'),   'true',  '0.5 → True');
    assert.strictEqual(runFunc(branch('-0.5'), 'Check'),  'true',  '-0.5 → True');
    console.log('[PASS] 数値の truthy 判定');
}

// =============================================================================
// 2. Boolean リテラルの truthy 判定
// =============================================================================
{
    assert.strictEqual(runFunc(branch('True'), 'Check'),  'true',  'True → True');
    assert.strictEqual(runFunc(branch('False'), 'Check'), 'false', 'False → False');
    console.log('[PASS] Boolean リテラルの truthy 判定');
}

// =============================================================================
// 3. 文字列の truthy 判定（VBA 仕様準拠）
// =============================================================================
{
    // "True"/"False" → そのまま Boolean に変換
    assert.strictEqual(runFunc(branch('"True"'), 'Check'),   'true',  '"True" → True');
    assert.strictEqual(runFunc(branch('"False"'), 'Check'),  'false', '"False" → False');
    assert.strictEqual(runFunc(branch('"true"'), 'Check'),   'true',  '"true" → True（大文字小文字無視）');
    assert.strictEqual(runFunc(branch('"FALSE"'), 'Check'),  'false', '"FALSE" → False（大文字小文字無視）');

    // 数値文字列 → 数値化して評価
    assert.strictEqual(runFunc(branch('"1"'), 'Check'),   'true',  '"1" → True');
    assert.strictEqual(runFunc(branch('"0"'), 'Check'),   'false', '"0" → False（数値 0）');
    assert.strictEqual(runFunc(branch('"-1"'), 'Check'),  'true',  '"-1" → True（数値 -1）');
    assert.strictEqual(runFunc(branch('"2"'), 'Check'),   'true',  '"2" → True');

    // 変換不能な文字列 → Type mismatch (Error 13)
    const ev1 = (() => {
        const tokens = new Lexer(branch('""')).tokenize();
        const ast = new Parser(tokens).parse();
        const ev = new Evaluator(console.log);
        ev.evaluate(ast);
        return ev;
    })();
    expectError(() => ev1.callProcedure('Check', []), 13, '"" → Type mismatch');

    const ev2 = (() => {
        const tokens = new Lexer(branch('"abc"')).tokenize();
        const ast = new Parser(tokens).parse();
        const ev = new Evaluator(console.log);
        ev.evaluate(ast);
        return ev;
    })();
    expectError(() => ev2.callProcedure('Check', []), 13, '"abc" → Type mismatch');

    console.log('[PASS] 文字列の truthy 判定');
}

// =============================================================================
// 4. Null → Error 94
// =============================================================================
{
    const code = `
        Function CheckNull() As String
            Dim v As Variant
            v = Null
            If v Then
                CheckNull = "true"
            Else
                CheckNull = "false"
            End If
        End Function
    `;
    const tokens = new Lexer(code).tokenize();
    const ast = new Parser(tokens).parse();
    const ev = new Evaluator(console.log);
    ev.evaluate(ast);
    expectError(() => ev.callProcedure('CheckNull', []), 94, 'Null → Invalid use of Null');
    console.log('[PASS] Null → Error 94');
}

// =============================================================================
// 5. Empty（未初期化 Variant）→ False
// =============================================================================
{
    const code = `
        Function CheckEmpty() As String
            Dim v As Variant   ' 未初期化 = Empty = 0
            If v Then
                CheckEmpty = "true"
            Else
                CheckEmpty = "false"
            End If
        End Function
    `;
    assert.strictEqual(runFunc(code, 'CheckEmpty'), 'false', 'Empty → False');
    console.log('[PASS] Empty → False');
}

// =============================================================================
// 6. Boolean 型の演算・変換
// =============================================================================
{
    // CInt(True) = -1、CInt(False) = 0
    assert.strictEqual(runFunc('Function F() As Long: F = CInt(True): End Function',  'F'), -1, 'CInt(True) = -1');
    assert.strictEqual(runFunc('Function F() As Long: F = CInt(False): End Function', 'F'),  0, 'CInt(False) = 0');

    // Boolean 算術（True = -1 として計算）
    assert.strictEqual(runFunc('Function F() As Long: F = True + 1:  End Function', 'F'),  0, 'True + 1 = 0（-1+1）');
    assert.strictEqual(runFunc('Function F() As Long: F = False + 1: End Function', 'F'),  1, 'False + 1 = 1');
    assert.strictEqual(runFunc('Function F() As Long: F = True * 2:  End Function', 'F'), -2, 'True * 2 = -2（-1*2）');

    // Not はビット反転
    assert.isTrue( runFunc('Function F() As Boolean: F = Not False: End Function', 'F'), 'Not False → True');
    assert.isFalse(runFunc('Function F() As Boolean: F = Not True:  End Function', 'F'), 'Not True → False');
    assert.strictEqual(runFunc('Function F() As Long: F = Not 0: End Function', 'F'), -1, 'Not 0 = -1（ビット反転）');
    assert.strictEqual(runFunc('Function F() As Long: F = Not 1: End Function', 'F'), -2, 'Not 1 = -2（ビット反転）');

    // CStr(True) = "True"、CStr(False) = "False"
    assert.strictEqual(runFunc('Function F() As String: F = CStr(True):  End Function', 'F'), 'True',  'CStr(True) = "True"');
    assert.strictEqual(runFunc('Function F() As String: F = CStr(False): End Function', 'F'), 'False', 'CStr(False) = "False"');

    console.log('[PASS] Boolean 型の演算・変換');
}

// =============================================================================
// 7. 比較演算子の結果は vbaTrue(-1) / vbaFalse(0)
// =============================================================================
{
    assert.isTrue( runFunc('Function F() As Boolean: F = (1 > 0): End Function',  'F'), '1 > 0 → vbaTrue');
    assert.isFalse(runFunc('Function F() As Boolean: F = (1 < 0): End Function',  'F'), '1 < 0 → vbaFalse');
    assert.strictEqual(
        runFunc('Function F() As Long: F = (1 > 0): End Function', 'F'), -1,
        '(1 > 0) を Long に代入 = -1'
    );
    assert.strictEqual(
        runFunc('Function F() As Long: F = (1 < 0): End Function', 'F'), 0,
        '(1 < 0) を Long に代入 = 0'
    );
    console.log('[PASS] 比較演算子の結果');
}

console.log('\n✅ If 条件式の truthy 判定・Boolean 型強制: 全テスト通過');
