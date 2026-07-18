/**
 * 実 VBA 差分テスト（scripts/gen-diff-corpus.ts + Excel 実行結果との突き合わせ）で
 * 発見・修正したバグのリグレッションテスト。
 *
 * 2337 式のコーパスを実 Excel VBA と vba-runner の両方で実行し、TypeName/値/
 * Err.Number まで比較した結果、2302/2337 が一致（残り 35 件はコードページ・
 * タイムゾーン等の環境依存差として scripts/diff-allowlist.txt に許容登録）。
 * ここでは発見された仕様差のうち代表的なものを固定する。
 */
import { evalVBASingle, assert } from '../../test-libs/test-runner';

const ev = evalVBASingle('');

// --- 括弧は数値サブタイプに影響しない ---
{
    assert.strictEqual(ev.evalExpression('TypeName((2))'), 'Integer', '(2) は Integer');
    assert.strictEqual(ev.evalExpression('TypeName((2) + (2))'), 'Integer', '(2)+(2) は Integer');
}

// --- Integer/Long 演算のオーバーフロー（型昇格せず Error 6） ---
{
    assert.throwsMatch(() => ev.evalExpression('(32767) + (2)'), /error '6'/, 'Integer + Integer オーバーフロー');
    assert.throwsMatch(() => ev.evalExpression('(2) + (32767)'), /error '6'/, 'オペランドの順序を入れ替えても同様');
    const dimmed = evalVBASingle('Function F() As Long\nDim a As Long\na = 32767\nF = a + 2\nEnd Function');
    assert.strictEqual(dimmed.callProcedure('F', []), 32769, 'Long 型変数が絡めば昇格して正常');
}

// --- 0 除算: 0/0 は Overflow(6)、非ゼロ/0 は Division by zero(11) ---
{
    assert.throwsMatch(() => ev.evalExpression('(0) / (0)'), /error '6'/, '0/0 は Error 6');
    assert.throwsMatch(() => ev.evalExpression('(1) / (0)'), /error '11'/, '1/0 は Error 11');
}

// --- CInt/CLng の丸めは Round 相当（銀行家丸め）だが結果の型タグは変換先型 ---
{
    assert.strictEqual(ev.evalExpression('TypeName(CInt(2.5))'), 'Integer', 'CInt の戻り値は Integer 型');
    assert.strictEqual(ev.evalExpression('CInt(2.5)'), 2, 'CInt(2.5) はバンカーズ丸めで 2');
}

// --- Format() の丸めは Round() と異なり通常の四捨五入（銀行家丸めではない） ---
{
    assert.strictEqual(ev.evalExpression('Format(0.5, "0")'), '1', 'Format(0.5,"0") は四捨五入で "1"');
    assert.strictEqual(ev.evalExpression('Format(-0.5, "0")'), '-1', 'Format(-0.5,"0") も四捨五入');
}

// --- CStr(Null) は Error 94（空文字列ではない） ---
{
    assert.throwsMatch(() => ev.evalExpression('CStr(Null)'), /error '94'/, 'CStr(Null) → Error 94');
}

// --- CDate(Empty) はシリアル値 0（1899-12-30） ---
{
    assert.strictEqual(ev.evalExpression('CStr(CDate(Empty))'), '1899/12/30', 'CDate(Empty) はエポック日');
}

// --- CDate: 数値文字列（&H/&O/カンマ）はシリアル値として解釈する ---
{
    assert.strictEqual(ev.evalExpression('CLng(CDate("&H10"))'), 16, 'CDate("&H10") は数値 16 のシリアル値');
}

// --- 文字列 op Null の Null 伝播は + のみ特別扱い、-/*/ は数値変換エラーが先行 ---
{
    assert.strictEqual(ev.evalExpression('IsNull(("abc") + (Null))'), ev.evalExpression('True'),
        '"abc" + Null は Null のまま伝播');
    assert.throwsMatch(() => ev.evalExpression('("abc") - (Null)'), /error '13'/,
        '"abc" - Null は数値変換失敗で Error 13（Null 伝播より先）');
}

// --- 比較演算子: 文字列と日付/数値の混在は数値変換で比較。変換不能なら Error 13 ---
{
    assert.throwsMatch(() => ev.evalExpression('("abc") = (#2024/03/15#)'), /error '13'/,
        '"abc" = #date# は Error 13');
    assert.throwsMatch(() => ev.evalExpression('(2) < ("abc")'), /error '13'/,
        '(2) < ("abc") も数値変換失敗で Error 13');
    assert.strictEqual(String(ev.evalExpression('(2) < ("7")')), 'True',
        '変換可能な数値文字列との比較は正常に動作する');
}

// --- 比較演算子: 片方が Boolean で片方が文字列なら CBool 変換して比較する ---
{
    assert.strictEqual(String(ev.evalExpression('("7") = (True)')), 'True',
        '"7" は CBool で True になるため True = True');
    assert.strictEqual(String(ev.evalExpression('("0") = (True)')), 'False',
        '"0" は CBool で False になるため False = True → False');
}

// --- Date ± Number は Date 型を維持し、範囲外は Overflow ---
{
    assert.strictEqual(ev.evalExpression('TypeName((3.5) - (#2024/03/15#))'), 'Date',
        '数値 - Date は Date 型（VBA は日付演算として扱う）');
    assert.strictEqual(ev.evalExpression('TypeName((#2024/03/15#) - (#2024/03/10#))'), 'Double',
        'Date - Date は数値');
}

// --- DateSerial: 月/日の範囲外は繰り上げ・繰り下げされる（年をまたぐケース） ---
{
    assert.strictEqual(ev.evalExpression('Year(DateSerial(2024, 13, 1))'), 2025,
        'DateSerial(2024,13,1) は月が繰り上がり 2025年');
    assert.strictEqual(ev.evalExpression('Year(DateSerial(2024, 0, 15))'), 2023,
        'DateSerial(2024,0,15) は月が繰り下がり 2023年');
}

// --- Sgn/Abs 等、片方の型が静的に不明な乗算はオーバーフロー誤検出しない ---
{
    // Sgn() は Integer 型固定、Abs() は引数依存で型不明。誤って Integer 昇格すると
    // 32767 を超える結果で Overflow の誤検出が起きる（実際に一度リグレッションした）
    const r = ev.evalExpression('Sgn(-500000.5) * Abs(-500000.5)');
    assert.strictEqual(Math.round(Number(r) * 10) / 10, -500000.5,
        'Sgn(x)*Abs(x) はオーバーフロー誤検出せず x を復元する');
}

// --- \ / Mod: 左が Boolean で右が文字列のときだけ CBool 変換し結果も Boolean 型になる
// （左右逆・両方 Boolean literal のときは通常の整数演算。非対称な規則を深掘りで発見） ---
{
    assert.strictEqual(String(ev.evalExpression('(True) \\ ("7")')), 'True',
        'True \\ "7": "7"→CBool True(-1)、True\\True=1→非ゼロで Boolean True');
    assert.strictEqual(String(ev.evalExpression('(False) \\ ("7")')), 'False',
        'False \\ "7": 0\\True(-1)=0 → Boolean False');
    assert.strictEqual(String(ev.evalExpression('(True) Mod ("7")')), 'False',
        'True Mod "7": -1 Mod -1=0 → Boolean False');
    assert.strictEqual(ev.evalExpression('("7") \\ (True)'), -7,
        '"7" \\ True: 左が文字列なので通常の数値変換。7\\(-1) = Long -7（Boolean にならない）');
    assert.strictEqual(ev.evalExpression('(True) \\ (True)'), 1,
        'True \\ True: 両方すでに Boolean literal（文字列変換なし）なので通常の整数昇格 Integer 1');
    console.log('[PASS] \\/Mod の Boolean 型保持は左 Boolean・右文字列のときだけの非対称規則');
}

// --- CDate/日付演算の "H.N" 文字列は時刻（H時N分）として解釈される（+ 演算子限定） ---
{
    assert.strictEqual(ev.evalExpression('CStr(CDate("2.5"))'), '02:05:00',
        'CDate("2.5") は "2時5分" と解釈される（シリアル値 2.5 ではない）');
    assert.strictEqual(ev.evalExpression('CStr((#2024/03/15#) + ("3.5"))'), '2024/03/15 03:05:00',
        '#date# + "3.5" は 03:05:00 を加算する（+ 演算子だけの特殊ルール）');
    assert.strictEqual(ev.evalExpression('CStr(("3.5") + (#2024/03/15#))'), '2024/03/15 03:05:00',
        'オペランドの順序を入れ替えても同じ（左右対称）');
    // - / * / \ / Mod では同じ文字列でも通常どおりシリアル値として解釈される（+ の非対称ルール）
    assert.strictEqual(ev.evalExpression('CLng((#2024/03/15#) * ("3.5"))'), 158781,
        '#date# * "3.5" は "3.5" をシリアル値として掛ける（時刻解釈はしない）');
    assert.strictEqual(String(ev.evalExpression('CStr((#2024/03/15#) - ("3.5"))')), '2024/03/11 12:00:00',
        '#date# - "3.5" も同様にシリアル値として減算する');
    console.log('[PASS] "H.N" 文字列の時刻解釈は + 演算子・CDate 限定（-/*/\\/Mod は通常の数値解釈）');
}

console.log('\n✅ real-vba-diff-regressions: 全テスト通過');
