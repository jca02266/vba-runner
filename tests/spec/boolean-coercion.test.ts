/**
 * Boolean の代入時の型強制 (Let-coercion §5.6.10 + §6.1.2.3.1.1)
 *
 * VBA の `Dim x As Boolean` で宣言された変数に代入する際、右辺の値を
 * Boolean に強制変換する（Let-coercion）。
 *
 *   - 数値: 0 → False (vbaFalse)、非 0 → True (vbaTrue)
 *   - 文字列 "True" / "False" (大文字小文字無視): 対応する Boolean
 *   - 文字列で数値表現: 数値化してから Boolean 変換
 *   - 変換不能な文字列: Type mismatch (Error 13)
 *   - Empty: False
 *   - Null: Invalid use of Null (Error 94)
 *
 * これは `assert.isTrue` などのリファレンス比較を前提に、シングルトン
 * `vbaTrue` / `vbaFalse` を返す必要がある。
 */
import { Lexer } from '../../src/compiler/lexer';
import { Parser } from '../../src/compiler/parser';
import { Evaluator, vbaTrue, vbaFalse } from '../../src/compiler/evaluator';
import { assert } from '../../test-libs/test-runner';

function evalVBA(code: string): any {
    const tokens = new Lexer(code).tokenize();
    const ast = new Parser(tokens).parse();
    const ev = new Evaluator(console.log);
    ev.evaluate(ast);
    return ev;
}

function runFunc(code: string, name: string, args: any[] = []): any {
    return evalVBA(code).callProcedure(name, args);
}

function expectError(fn: () => any, errorNumber: number, label: string) {
    let caught: any = null;
    try { fn(); } catch (e: any) { caught = e; }
    if (!caught) {
        console.error(`[FAIL] ${label} - エラーが発生しなかった`);
        throw new Error('Assertion Failed');
    }
    if (Number(caught.number) !== errorNumber) {
        console.error(`[FAIL] ${label} - 期待エラー番号 ${errorNumber}、実際: ${JSON.stringify(caught)}`);
        throw new Error('Assertion Failed');
    }
}

console.log('--- Starting Boolean Coercion Tests ---');

// =============================================================================
// 1. 数値からの Boolean 変換
// =============================================================================
{
    const code = `
        Function FromInt(n As Long) As Boolean
            Dim b As Boolean
            b = n
            FromInt = b
        End Function

        Function FromDouble(d As Double) As Boolean
            Dim b As Boolean
            b = d
            FromDouble = b
        End Function
    `;
    assert.isTrue(runFunc(code, 'FromInt', [5]),     '数値 5 → True');
    assert.isFalse(runFunc(code, 'FromInt', [0]),    '数値 0 → False');
    assert.isTrue(runFunc(code, 'FromInt', [-1]),    '数値 -1 → True');
    assert.isTrue(runFunc(code, 'FromInt', [-100]),  '負の数値 → True');
    assert.isTrue(runFunc(code, 'FromDouble', [0.5]),  '小数 0.5 → True');
    assert.isFalse(runFunc(code, 'FromDouble', [0.0]), '小数 0.0 → False');
    console.log('[PASS] 数値からの Boolean 変換');
}

// =============================================================================
// 2. 文字列 "True"/"False" からの Boolean 変換（大文字小文字無視）
// =============================================================================
{
    const code = `
        Function FromStr(s As String) As Boolean
            Dim b As Boolean
            b = s
            FromStr = b
        End Function
    `;
    assert.isTrue(runFunc(code, 'FromStr', ['True']),    '"True" → True');
    assert.isFalse(runFunc(code, 'FromStr', ['False']),  '"False" → False');
    assert.isTrue(runFunc(code, 'FromStr', ['true']),    '"true" → True');
    assert.isFalse(runFunc(code, 'FromStr', ['FALSE']),  '"FALSE" → False');
    assert.isTrue(runFunc(code, 'FromStr', ['TRUE']),    '"TRUE" → True');
    console.log('[PASS] 文字列 "True"/"False" からの変換');
}

// =============================================================================
// 3. 数値文字列からの Boolean 変換
// =============================================================================
{
    const code = `
        Function FromNumStr(s As String) As Boolean
            Dim b As Boolean
            b = s
            FromNumStr = b
        End Function
    `;
    assert.isTrue(runFunc(code, 'FromNumStr', ['1']),    '"1" → True');
    assert.isFalse(runFunc(code, 'FromNumStr', ['0']),   '"0" → False');
    assert.isTrue(runFunc(code, 'FromNumStr', ['-1']),   '"-1" → True');
    assert.isTrue(runFunc(code, 'FromNumStr', ['100']),  '"100" → True');
    console.log('[PASS] 数値文字列からの変換');
}

// =============================================================================
// 4. Empty からの Boolean 変換
// =============================================================================
{
    const code = `
        Function FromEmpty() As Boolean
            Dim v As Variant   ' 未初期化 → Empty
            Dim b As Boolean
            b = v
            FromEmpty = b
        End Function
    `;
    assert.isFalse(runFunc(code, 'FromEmpty'), 'Empty → False');
    console.log('[PASS] Empty からの Boolean 変換');
}

// =============================================================================
// 5. エラーケース
// =============================================================================
{
    const code = `
        Function FromBadStr() As Boolean
            Dim b As Boolean
            b = "abc"
            FromBadStr = b
        End Function

        Function FromNull() As Boolean
            Dim b As Boolean
            b = Null
            FromNull = b
        End Function
    `;
    const ev = evalVBA(code);
    expectError(() => ev.callProcedure('FromBadStr', []), 13, '"abc" → Boolean は Type mismatch');
    expectError(() => ev.callProcedure('FromNull', []), 94, 'Null → Boolean は Invalid use of Null');
    console.log('[PASS] 変換不能な値でのエラー');
}

// =============================================================================
// 6. VBA ベテランが書きそうな代表コード — フラグ管理
// =============================================================================
{
    const code = `
        ' レジストリやセル値（文字列で来る）を Boolean に変換する典型パターン
        Function ParseSetting(rawValue As String) As Boolean
            Dim enabled As Boolean
            enabled = rawValue   ' "True"/"False" を直接代入
            ParseSetting = enabled
        End Function

        ' 整数フラグ（0/1）を Boolean に変換する典型パターン
        Function IsEnabled(intFlag As Long) As Boolean
            Dim enabled As Boolean
            enabled = intFlag    ' 0 → False, 非 0 → True
            IsEnabled = enabled
        End Function

        ' 三項条件式の代用（IIf 経由）
        Function Threshold(score As Long, target As Long) As Boolean
            Dim pass As Boolean
            pass = (score >= target)  ' 比較結果を直接代入
            Threshold = pass
        End Function
    `;
    assert.isTrue(runFunc(code, 'ParseSetting', ['True']),    '設定 "True" → 有効');
    assert.isFalse(runFunc(code, 'ParseSetting', ['False']),  '設定 "False" → 無効');
    assert.isTrue(runFunc(code, 'IsEnabled', [1]),    'フラグ 1 → 有効');
    assert.isFalse(runFunc(code, 'IsEnabled', [0]),   'フラグ 0 → 無効');
    assert.isTrue(runFunc(code, 'Threshold', [80, 70]),   '80 >= 70 → True');
    assert.isFalse(runFunc(code, 'Threshold', [60, 70]),  '60 < 70 → False');
    console.log('[PASS] VBA ベテランパターン: フラグ管理');
}

// =============================================================================
// 7. 既に Boolean の場合: シングルトン同等性が保たれる
// =============================================================================
{
    const code = `
        Function PassThrough(input As Boolean) As Boolean
            Dim b As Boolean
            b = input
            PassThrough = b
        End Function
    `;
    // 入力が vbaTrue なら戻り値も vbaTrue（シングルトン）
    const t = runFunc(code, 'PassThrough', [vbaTrue]);
    const f = runFunc(code, 'PassThrough', [vbaFalse]);
    assert.strictEqual(t, vbaTrue, 'vbaTrue 入出力でシングルトン同一');
    assert.strictEqual(f, vbaFalse, 'vbaFalse 入出力でシングルトン同一');
    console.log('[PASS] Boolean シングルトン保持');
}

console.log('\n✅ Boolean 代入時の型強制: 全テスト通過');
