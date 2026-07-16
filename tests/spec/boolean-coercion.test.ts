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
import { vbaTrue, vbaFalse, vbaNull } from '../../src/engine/evaluator';
import { evalVBASingle, assert } from '../../test-libs/test-runner';

function evalVBA(code: string): any {
    return evalVBASingle(code);
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
    assert.strictEqual(runFunc(code, 'FromInt', [5]), vbaTrue, '数値 5 → True');
    assert.strictEqual(runFunc(code, 'FromInt', [0]), vbaFalse, '数値 0 → False');
    assert.strictEqual(runFunc(code, 'FromInt', [-1]), vbaTrue, '数値 -1 → True');
    assert.strictEqual(runFunc(code, 'FromInt', [-100]), vbaTrue, '負の数値 → True');
    assert.strictEqual(runFunc(code, 'FromDouble', [0.5]), vbaTrue, '小数 0.5 → True');
    assert.strictEqual(runFunc(code, 'FromDouble', [0.0]), vbaFalse, '小数 0.0 → False');
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
    assert.strictEqual(runFunc(code, 'FromStr', ['True']), vbaTrue, '"True" → True');
    assert.strictEqual(runFunc(code, 'FromStr', ['False']), vbaFalse, '"False" → False');
    assert.strictEqual(runFunc(code, 'FromStr', ['true']), vbaTrue, '"true" → True');
    assert.strictEqual(runFunc(code, 'FromStr', ['FALSE']), vbaFalse, '"FALSE" → False');
    assert.strictEqual(runFunc(code, 'FromStr', ['TRUE']), vbaTrue, '"TRUE" → True');
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
    assert.strictEqual(runFunc(code, 'FromNumStr', ['1']), vbaTrue, '"1" → True');
    assert.strictEqual(runFunc(code, 'FromNumStr', ['0']), vbaFalse, '"0" → False');
    assert.strictEqual(runFunc(code, 'FromNumStr', ['-1']), vbaTrue, '"-1" → True');
    assert.strictEqual(runFunc(code, 'FromNumStr', ['100']), vbaTrue, '"100" → True');
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
    assert.strictEqual(runFunc(code, 'FromEmpty'), vbaFalse, 'Empty → False');
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
    assert.strictEqual(runFunc(code, 'ParseSetting', ['True']), vbaTrue, '設定 "True" → 有効');
    assert.strictEqual(runFunc(code, 'ParseSetting', ['False']), vbaFalse, '設定 "False" → 無効');
    assert.strictEqual(runFunc(code, 'IsEnabled', [1]), vbaTrue, 'フラグ 1 → 有効');
    assert.strictEqual(runFunc(code, 'IsEnabled', [0]), vbaFalse, 'フラグ 0 → 無効');
    assert.strictEqual(runFunc(code, 'Threshold', [80, 70]), vbaTrue, '80 >= 70 → True');
    assert.strictEqual(runFunc(code, 'Threshold', [60, 70]), vbaFalse, '60 < 70 → False');
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

// --- Bug BO: evalExpression("True + True") が Error 424 になっていた ---
// isCallableLeftmostLeaf が True/False 等の定数 Identifier も callable とみなし、
// statement fallback に飛ばしていたため Object required エラーになっていた。
{
    const ev = evalVBASingle('');
    assert.strictEqual(ev.evalExpression('True + True'), -2, 'True + True = -2');
    assert.strictEqual(ev.evalExpression('True + False'), -1, 'True + False = -1');
    assert.strictEqual(ev.evalExpression('False + 1'), 1, 'False + 1 = 1');
    assert.strictEqual(ev.evalExpression('True + 1'), 0, 'True + 1 = 0');
    assert.strictEqual(ev.evalExpression('Null + 1'), vbaNull, 'Null + 1 = Null (propagation)');
    console.log('[PASS] Bug BO: True/False/Null in evalExpression arithmetic');
}

// --- Bug BP: evalExpression("\"abc\" = \"ABC\"") が undefined を返していた ---
// isStatementAmbiguous の = チェックが左辺の型に関係なくトリガーされ、
// 文字列リテラル = 文字列リテラルの比較が statement fallback に飛んで結果が捨てられていた。
// 修正: = が statement-ambiguous なのは左辺が代入可能 (Identifier 等) の場合のみ。
{
    const ev = evalVBASingle('');
    const isFalse = (v: any) => v && typeof v === 'object' && v.value === 0;
    const isTrue = (v: any) => v && typeof v === 'object' && v.value === -1;
    assert.ok(isFalse(ev.evalExpression('"abc" = "ABC"')), '"abc" = "ABC" → false');
    assert.ok(isTrue(ev.evalExpression('"abc" = "abc"')), '"abc" = "abc" → true');
    assert.ok(isTrue(ev.evalExpression('1 = 1')), '1 = 1 → true');
    assert.ok(isFalse(ev.evalExpression('1 = 2')), '1 = 2 → false');
    assert.strictEqual(ev.evalExpression('Null = 1'), vbaNull, 'Null = 1 → Null (propagation)');
    // Assignment still works (x is identifier → still ambiguous → statement path)
    const ev2 = evalVBASingle('Public x');
    ev2.evalExpression('x = 42');
    assert.strictEqual(ev2.env.get('x'), 42, 'x = 42 still executes as assignment');
    console.log('[PASS] Bug BP: literal comparison in evalExpression');
}

console.log('\n✅ Boolean 代入時の型強制: 全テスト通過');
