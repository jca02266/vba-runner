/**
 * Bug DT: scripts/fuzz-builtins.ts（組み込み関数の異常値スモークファザー）で検出した
 * クラッシュ群のリグレッションテスト。
 *
 * 修正前は以下がすべて JS の生例外（TypeError / RangeError / SyntaxError / generic Error）
 * で落ちていた。修正後は VBA エラー（Run-time error 'NN'）として制御される。
 *
 * 中央修正:
 * - coerce.ts vbaToNumber/vbaToString/vbaToBoolean: Nothing → Error 91、
 *   配列・非数値 → Error 13、undefined → 0（Empty 相当）
 * - vba-types.ts parseVbaDate: Nothing → Error 91、その他 Symbol → Error 13
 * - vba-types.ts VbaCurrency.fromNumber: 非有限・明確な範囲外 → Error 6
 *   （境界付近は従来どおり正確な文字列→BigInt 経路で判定し精度を保つ）
 * - vba-errors.ts: OUT_OF_STRING_SPACE (14) を追加
 */
import { evalVBASingle, assert } from '../../test-libs/test-runner';

const ev = evalVBASingle('');

// --- 1. Nothing の値コンテキスト使用 → Error 91 ---
{
    for (const expr of [
        'Abs(Nothing)', 'Sin(Nothing)', 'CInt(Nothing)', 'CDbl(Nothing)',
        'Trim(Nothing)', 'UCase(Nothing)', 'Space(Nothing)',
        'Year(Nothing)', 'Weekday(Nothing)', 'DateSerial(Nothing, 1, 1)',
        'FV(Nothing, 1, 1)', 'SLN(Nothing, 1, 1)',
        'Left("ab", Nothing)', 'Mid("ab", Nothing)', 'Round(1, Nothing)',
        'Chr(Nothing)', 'Choose(Nothing)', 'MonthName(Nothing)',
        'Format(Nothing)', 'FormatDateTime(Nothing)',
    ]) {
        assert.throwsMatch(() => ev.evalExpression(expr), /error '91'/, `${expr} → Error 91`);
    }
    console.log('[PASS] Nothing の値コンテキスト使用 → Error 91');
}

// --- 2. 配列・非数値の数値強制 → Error 13 ---
{
    assert.throwsMatch(() => ev.evalExpression('CInt(Array(1, 2, 3))'), /error '13'/, 'CInt(Array) → Error 13');
    assert.throwsMatch(() => ev.evalExpression('CLngLng(Array(1, 2, 3))'), /error '13'/, 'CLngLng(Array) → Error 13');
    assert.throwsMatch(() => ev.evalExpression('CCur(Array(1, 2, 3))'), /error '13'/, 'CCur(Array) → Error 13');
    console.log('[PASS] 配列の数値強制 → Error 13');
}

// --- 3. Space/String の Long 超過・巨大割り当て → Error 6 / 14 ---
{
    assert.throwsMatch(() => ev.evalExpression('Space(2147483648)'), /error '6'/, 'Space(Long超過) → Error 6');
    assert.throwsMatch(() => ev.evalExpression('String(2147483648, "x")'), /error '6'/, 'String(Long超過) → Error 6');
    // Long 範囲内だが JS の文字列上限（約 2^30）を超える → Error 14 Out of string space
    assert.throwsMatch(() => ev.evalExpression('Space(2147483647)'), /error '14'/, 'Space(Long最大) → Error 14');
    assert.strictEqual(ev.evalExpression('Space(3)'), '   ', 'Space(3) は正常動作');
    console.log('[PASS] Space/String の巨大割り当て → Error 6 / 14');
}

// --- 4. CCur のオーバーフローと境界精度 ---
{
    assert.throwsMatch(() => ev.evalExpression('CCur(1E308)'), /error '6'/, 'CCur(1E308) → Error 6');
    assert.throwsMatch(() => ev.evalExpression('CCur(1E300)'), /error '6'/, 'CCur(1E300) → Error 6');
    assert.throwsMatch(() => ev.evalExpression('CCur(922337203685478)'), /error '6'/, 'CCur(境界+1) → Error 6');
    // 境界値ちょうどは正確な文字列→BigInt 経路で受理される（double 経由の粗い比較で誤判定しない）
    assert.strictEqual(String(ev.evalExpression('CCur("922337203685477.5807")')), '922337203685477.5807',
        'CCur(Currency最大値文字列) は精度を失わず受理');
    console.log('[PASS] CCur のオーバーフローと境界精度');
}

// --- 5. FormatNumber 系の桁数境界 ---
{
    assert.throwsMatch(() => ev.evalExpression('FormatNumber(1, 2147483648)'), /error '5'/, 'FormatNumber(桁数>255) → Error 5');
    assert.throwsMatch(() => ev.evalExpression('FormatPercent(1, 300)'), /error '5'/, 'FormatPercent(桁数>255) → Error 5');
    // 101〜255 桁は 0 埋めで許容（VBA 上限 255、JS toFixed 上限 100 の差を吸収）
    const s = ev.evalExpression('FormatNumber(1.5, 110)') as string;
    assert.strictEqual(s.length, '1.'.length + 110, 'FormatNumber(1.5, 110) は 110 桁の小数部');
    console.log('[PASS] FormatNumber 系の桁数境界');
}

// --- 6. ファイル I/O のパス・ハンドル異常値 ---
{
    assert.throwsMatch(() => ev.evalExpression('FileLen("noexist_dt")'), /error '53'/, 'FileLen(不存在) → Error 53');
    assert.throwsMatch(() => ev.evalExpression('FileDateTime("noexist_dt")'), /error '53'/, 'FileDateTime(不存在) → Error 53');
    assert.throwsMatch(() => ev.evalExpression('FileCopy "noexist_dt", "dest_dt"'), /error '53'/, 'FileCopy(不存在) → Error 53');
    assert.throwsMatch(() => ev.evalExpression('EOF(Null)'), /error '13'/, 'EOF(Null) → Error 13');
    assert.throwsMatch(() => ev.evalExpression('Kill Null'), /error '94'/, 'Kill Null → Error 94');
    assert.throwsMatch(() => ev.evalExpression('MkDir Null'), /error '94'/, 'MkDir Null → Error 94');
    console.log('[PASS] ファイル I/O の異常値 → VBA エラー');
}

// --- 7. Interaction / Registry / COM 系の Null・Nothing ---
{
    assert.throwsMatch(() => ev.evalExpression('MsgBox(Null)'), /error '94'/, 'MsgBox(Null) → Error 94');
    assert.throwsMatch(() => ev.evalExpression('Shell(Null)'), /error '94'/, 'Shell(Null) → Error 94');
    assert.throwsMatch(() => ev.evalExpression('CreateObject(Null)'), /error '94'/, 'CreateObject(Null) → Error 94');
    // GetSetting: default が Null なら Error 94（従来は "Symbol(vbaNull)" が漏れていた）
    assert.throwsMatch(() => ev.evalExpression('GetSetting("a", "b", "c", Null)'), /error '94'/, 'GetSetting(default=Null) → Error 94');
    assert.strictEqual(ev.evalExpression('GetSetting("a", "b", "c")'), '', 'GetSetting 未登録キー → 既定値 ""');
    console.log('[PASS] Interaction/Registry/COM 系の Null・Nothing → VBA エラー');
}

console.log('\n✅ fuzz-regressions: 全テスト通過');
