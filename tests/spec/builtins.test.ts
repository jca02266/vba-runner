import { Lexer } from '../../src/engine/lexer';
import { Parser } from '../../src/engine/parser';
import { Evaluator, vbaTrue, vbaFalse, vbaNull } from '../../src/engine/evaluator';
import { assert, evalVBASingle } from '../../test-libs/test-runner';

function evalExpr(expr: string): any {
    const tokens = new Lexer(expr).tokenize();
    const parser = new Parser(tokens);
    const ast = (parser as any).parseExpression();
    const ev = new Evaluator(() => {});
    return (ev as any).evaluateExpression(ast);
}

// 1. String Functions
{
    assert.strictEqual(evalExpr('Len("Hello")'), 5, 'Len');
    assert.strictEqual(evalExpr('Left("Hello", 2)'), 'He', 'Left');
    assert.strictEqual(evalExpr('Right("Hello", 2)'), 'lo', 'Right');
    assert.strictEqual(evalExpr('Mid("Hello", 2, 2)'), 'el', 'Mid(start, len)');
    assert.strictEqual(evalExpr('Mid("Hello", 2)'), 'ello', 'Mid(start)');
    assert.strictEqual(evalExpr('InStr("Hello", "e")'), 2, 'InStr found');
    assert.strictEqual(evalExpr('InStr("Hello", "z")'), 0, 'InStr not found');
    assert.strictEqual(evalExpr('LCase("HELLO")'), 'hello', 'LCase');
    assert.strictEqual(evalExpr('Replace("Hello", "e", "a")'), 'Hallo', 'Replace');
}

// 2. Conversion Functions
{
    assert.strictEqual(evalExpr('CInt("123.5")'), 124, 'CInt (round)');
    assert.strictEqual(evalExpr('CStr(123)'), '123', 'CStr');
    assert.strictEqual(evalExpr('CBool(1)'), vbaTrue, 'CBool(1)');
    assert.strictEqual(evalExpr('1 = 1'), vbaTrue, 'Equality 1 = 1');
    assert.strictEqual(evalExpr('CBool(0)'), vbaFalse, 'CBool(0)');
    assert.strictEqual(evalExpr('Fix(123.9)'), 123, 'Fix positive');
    assert.strictEqual(evalExpr('Fix(-123.9)'), -123, 'Fix negative');
}

// 3. Math Functions
{
    assert.strictEqual(evalExpr('Abs(-10)'), 10, 'Abs');
    assert.strictEqual(evalExpr('Round(123.456, 2)'), 123.46, 'Round(n, m)');
    assert.strictEqual(evalExpr('Sqr(16)'), 4, 'Sqr');
}

// 4. Information Functions
{
    assert.strictEqual(evalExpr('IsNull(Null)'), vbaTrue, 'IsNull(Null)');
    assert.strictEqual(evalExpr('IsNull(123)'), vbaFalse, 'IsNull(123)');
    assert.strictEqual(evalExpr('IsArray(Array(1, 2))'), vbaTrue, 'IsArray(Array())');
    assert.strictEqual(evalExpr('IsArray(123)'), vbaFalse, 'IsArray(123)');
    assert.strictEqual(evalExpr('IsObject(CreateObject("Scripting.Dictionary"))'), vbaTrue, 'IsObject(Dictionary)');
    assert.strictEqual(evalExpr('IsObject(123)'), vbaFalse, 'IsObject(123)');
}

// 5. Array Functions
{
    assert.strictEqual(evalExpr('LBound(Array(1, 2, 3))'), 0, 'LBound');
    assert.strictEqual(evalExpr('Array(1, 2, 3)(0)'), 1, 'Array index 0');
    assert.strictEqual(evalExpr('Array(1, 2, 3)(2)'), 3, 'Array index 2');
}

// 6. IIf
{
    assert.strictEqual(evalExpr('IIf(True, "A", "B")'), 'A', 'IIf True');
    assert.strictEqual(evalExpr('IIf(False, "A", "B")'), 'B', 'IIf False');
}

// RGB / QBColor / Nz
{
    assert.strictEqual(evalExpr('RGB(255, 0, 0)'), 255, 'RGB red');
    assert.strictEqual(evalExpr('RGB(0, 255, 0)'), 65280, 'RGB green');
    assert.strictEqual(evalExpr('RGB(0, 0, 255)'), 16711680, 'RGB blue');
    assert.strictEqual(evalExpr('RGB(0, 0, 0)'), 0, 'RGB black');
    assert.strictEqual(evalExpr('RGB(255, 255, 255)'), 16777215, 'RGB white');
    assert.strictEqual(evalExpr('QBColor(0)'), 0, 'QBColor(0) = black');
    assert.strictEqual(evalExpr('QBColor(15)'), 16777215, 'QBColor(15) = white');
    assert.strictEqual(evalExpr('QBColor(4)'), 128, 'QBColor(4) = dark red');
    assert.strictEqual(evalExpr('Nz(Null, 99)'), 99, 'Nz(Null, 99) = 99');
    assert.strictEqual(evalExpr('Nz(42, 99)'), 42, 'Nz(42, 99) = 42');
    assert.strictEqual(evalExpr('Nz(Null)'), '', 'Nz(Null) default = "" (VBA default is empty string, not 0)');
    console.log('[PASS] RGB / QBColor / Nz');
}

// Format named boolean formats / StrConv LCID / CreateObject ServerName / InputBox+MsgBox extra params
{
    // Format "Yes/No", "On/Off", "True/False" with VbaBoolean
    assert.strictEqual(evalExpr('Format(True, "Yes/No")'), 'Yes', 'Format True Yes/No');
    assert.strictEqual(evalExpr('Format(False, "Yes/No")'), 'No', 'Format False Yes/No');
    assert.strictEqual(evalExpr('Format(True, "On/Off")'), 'On', 'Format True On/Off');
    assert.strictEqual(evalExpr('Format(False, "On/Off")'), 'Off', 'Format False On/Off');
    assert.strictEqual(evalExpr('Format(True, "True/False")'), 'True', 'Format True True/False');
    assert.strictEqual(evalExpr('Format(False, "True/False")'), 'False', 'Format False True/False');
    // StrConv with LCID (3rd arg)
    assert.strictEqual(evalExpr('StrConv("hello", 1, 1033)'), 'HELLO', 'StrConv 3-arg LCID');
    console.log('[PASS] Format named boolean / StrConv LCID');
}

// Bug N2: Format("Scientific") が小文字 e・1桁指数を返す
{
    assert.strictEqual(evalExpr('Format(1234567, "Scientific")'), '1.23E+06', 'Format Scientific 大文字E・2桁指数');
    assert.strictEqual(evalExpr('Format(0.00123, "Scientific")'), '1.23E-03', 'Format Scientific 負指数');
    assert.strictEqual(evalExpr('Format(1, "Scientific")'), '1.00E+00', 'Format Scientific 指数0');
    console.log('[PASS] Bug N2: Format("Scientific") E+NN 形式');
}

// Bug L: Format(VbaCurrency/VbaDecimal, 数値パターン) が書式適用されない
{
    assert.strictEqual(evalExpr('Format(CCur(1234.5), "#,##0.00")'), '1,234.50', 'Format(CCur, "#,##0.00")');
    assert.strictEqual(evalExpr('Format(CCur(0), "0.00")'), '0.00', 'Format(CCur(0), "0.00")');
    console.log('[PASS] Bug L: Format(Currency/Decimal, 数値パターン)');
}

// Bug F: Format(VbaBoolean, 数値パターン)
{
    assert.strictEqual(evalExpr('Format(True, "0")'), '-1', 'Format(True, "0") = "-1"');
    assert.strictEqual(evalExpr('Format(False, "0")'), '0', 'Format(False, "0") = "0"');
    assert.strictEqual(evalExpr('Format(True, "#")'), '-1', 'Format(True, "#")');
    assert.strictEqual(evalExpr('Format(True)'), 'True', 'Format(True) 書式なし = "True"');
    console.log('[PASS] Bug F: Format(Boolean, 数値パターン)');
}

// Bug G: IsDate(数値) = True
{
    assert.strictEqual(evalExpr('IsDate(1)'), vbaTrue, 'IsDate(1) = True');
    assert.strictEqual(evalExpr('IsDate(0)'), vbaTrue, 'IsDate(0) = True');
    assert.strictEqual(evalExpr('IsDate(-1)'), vbaTrue, 'IsDate(-1) = True');
    assert.strictEqual(evalExpr('IsDate("abc")'), vbaFalse, 'IsDate("abc") = False');
    console.log('[PASS] Bug G: IsDate(数値) = True');
}

// Bug Y: Format(文字列, 日付パターン) が文字列フォーマットとして処理される問題
{
    const runFmt = (d: string, fmt: string) => {
        const r = evalVBASingle(`Function F() : F = Format(CDate("${d}"), "${fmt}") : End Function`);
        return r.callProcedure('F', []);
    };
    assert.strictEqual(runFmt('2024/03/15', 'yyyy'), '2024',  'Format(CDate, "yyyy") = "2024"');
    assert.strictEqual(runFmt('2024/03/15', 'mm'),   '03',    'Format(CDate, "mm") = "03"');
    assert.strictEqual(runFmt('2024/03/15', 'dd'),   '15',    'Format(CDate, "dd") = "15"');
    assert.strictEqual(runFmt('2024/03/15', 'mmmm'), 'March', 'Format(CDate, "mmmm") = "March"');
    // String input with date pattern should also work
    assert.strictEqual(evalExpr('Format("2024/03/15", "yyyy")'), '2024', 'Format(string, "yyyy") = "2024"');
    assert.strictEqual(evalExpr('Format("2024/03/15", "mm")'),   '03',   'Format(string, "mm") = "03"');
    console.log('[PASS] Bug Y: Format(文字列, 日付パターン)');
}

// Bug Z: AM/PM 書式指定時に時刻が 24 時間制のままになる問題
{
    const fmtDate = (code: string) =>
        evalVBASingle(`Function F(): F = ${code} : End Function`).callProcedure('F', []);
    assert.strictEqual(fmtDate('Format(CDate("14:30:00"), "hh:nn AM/PM")'), '02:30 PM', 'hh:nn AM/PM 14:30 → 02:30 PM');
    assert.strictEqual(fmtDate('Format(CDate("09:30:00"), "hh:nn am/pm")'), '09:30 am', 'hh:nn am/pm 09:30 → 09:30 am');
    assert.strictEqual(fmtDate('Format(CDate("00:30:00"), "h:nn AM/PM")'),  '12:30 AM', 'h:nn AM/PM 00:30 → 12:30 AM');
    assert.strictEqual(fmtDate('Format(CDate("12:00:00"), "h:nn AM/PM")'),  '12:00 PM', 'h:nn AM/PM 12:00 → 12:00 PM');
    assert.strictEqual(fmtDate('Format(CDate("23:59:59"), "hh:nn:ss AM/PM")'), '11:59:59 PM', 'hh:nn:ss AM/PM 23:59:59 → 11:59:59 PM');
    assert.strictEqual(fmtDate('Format(CDate("14:30:00"), "hh:nn")'), '14:30', 'hh:nn without AM/PM → 24h 14:30');
    console.log('[PASS] Bug Z: AM/PM 書式指定時の 12 時間制変換');
}

// Bug BC: FormatCurrency / FormatNumber / FormatPercent / FormatDateTime が未実装
{
    assert.strictEqual(evalExpr('FormatCurrency(1234.5)'), '$1,234.50', 'FormatCurrency(1234.5)');
    assert.strictEqual(evalExpr('FormatCurrency(-1234.5)'), '-$1,234.50', 'FormatCurrency(-1234.5)');
    assert.strictEqual(evalExpr('FormatCurrency(1234.5, 0)'), '$1,235', 'FormatCurrency(n, 0 digits)');
    assert.strictEqual(evalExpr('FormatNumber(1234.5)'), '1,234.50', 'FormatNumber(1234.5)');
    assert.strictEqual(evalExpr('FormatNumber(0.123, 4)'), '0.1230', 'FormatNumber(n, 4 digits)');
    assert.strictEqual(evalExpr('FormatPercent(0.5)'), '50.00%', 'FormatPercent(0.5)');
    assert.strictEqual(evalExpr('FormatPercent(0.1234, 1)'), '12.3%', 'FormatPercent(0.1234, 1)');
    assert.strictEqual(evalExpr('FormatCurrency(Null)'), '', 'FormatCurrency(Null) = ""');
    const fmtDt = (code: string) =>
        evalVBASingle(`Function F(): F = ${code}: End Function`).callProcedure('F', []);
    assert.strictEqual(fmtDt('FormatDateTime(CDate("2024/3/15"))'), '3/15/2024', 'FormatDateTime date-only');
    assert.strictEqual(fmtDt('FormatDateTime(CDate("2024/3/15"), 2)'), '3/15/2024', 'FormatDateTime vbShortDate');
    assert.strictEqual(fmtDt('FormatDateTime(CDate("14:30:00"), 4)'), '14:30', 'FormatDateTime vbShortTime');
    console.log('[PASS] Bug BC: FormatCurrency/FormatNumber/FormatPercent/FormatDateTime 実装');
}

// Bug AZ/BA/BB: VBA 三値論理 — And/Or/Imp が Null を含む場合の短絡評価
{
    // And: False が吸収元 → False And Null = False
    assert.strictEqual(evalExpr('False And Null') === vbaFalse, true, 'False And Null = False');
    assert.strictEqual(evalExpr('Null And False') === vbaFalse, true, 'Null And False = False');
    assert.strictEqual(evalExpr('True And Null') === vbaNull, true,  'True And Null = Null');
    // Or: True が吸収元 → True Or Null = True
    assert.strictEqual(evalExpr('True Or Null') === vbaTrue, true,   'True Or Null = True');
    assert.strictEqual(evalExpr('Null Or True') === vbaTrue, true,   'Null Or True = True');
    assert.strictEqual(evalExpr('False Or Null') === vbaNull, true,  'False Or Null = Null');
    // Imp: False→True, *→True、それ以外は Null
    assert.strictEqual(evalExpr('False Imp Null') === vbaTrue, true, 'False Imp Null = True');
    assert.strictEqual(evalExpr('Null Imp True') === vbaTrue, true,  'Null Imp True = True');
    assert.strictEqual(evalExpr('True Imp Null') === vbaNull, true,  'True Imp Null = Null');
    console.log('[PASS] Bug AZ/BA/BB: And/Or/Imp の三値論理 Null 伝播');
}

// Bug AO: Choose(Null, ...) が Type mismatch エラーを返す（VBA 仕様では Null を返す）
{
    assert.strictEqual(evalExpr('Choose(Null, "a", "b")') === vbaNull, true, 'Choose(Null, ...) = Null');
    assert.strictEqual(evalExpr('Choose(1, "a", "b")'), 'a', 'Choose(1, ...) = first choice');
    console.log('[PASS] Bug AO: Choose(Null) = Null');
}

// Bug AC/AD: Replace/Like の Null 伝播
{
    assert.strictEqual(evalExpr('Replace("abc", Null, "x")') === vbaNull, true, 'Replace(str,Null,repl) = Null');
    assert.strictEqual(evalExpr('Replace("abc", "a", Null)') === vbaNull, true, 'Replace(str,find,Null) = Null');
    assert.strictEqual(evalExpr('Null Like "*"') === vbaNull, true, 'Null Like pattern = Null');
    assert.strictEqual(evalExpr('"abc" Like Null') === vbaNull, true, 'str Like Null = Null');
    console.log('[PASS] Bug AC/AD: Replace/Like の Null 伝播');
}

// Bug AE: Chr(Null) / ChrW(Null) が TypeError でクラッシュする
{
    assert.strictEqual(evalExpr('Chr(Null)') === vbaNull, true, 'Chr(Null) = Null');
    assert.strictEqual(evalExpr('ChrW(Null)') === vbaNull, true, 'ChrW(Null) = Null');
    console.log('[PASS] Bug AE: Chr/ChrW の Null 伝播');
}

// Bug AF: Hex/Oct が Math.floor で切り捨てていた（VBA では四捨五入）
{
    assert.strictEqual(evalExpr('Hex(3.7)'), '4',  'Hex(3.7) = "4" (rounded)');
    assert.strictEqual(evalExpr('Hex(2.5)'), '2',  'Hex(2.5) = "2" (banker round)');
    assert.strictEqual(evalExpr('Hex(3.5)'), '4',  'Hex(3.5) = "4" (banker round)');
    assert.strictEqual(evalExpr('Oct(3.7)'), '4',  'Oct(3.7) = "4" (rounded)');
    console.log('[PASS] Bug AF: Hex/Oct の四捨五入');
}

// Bug BD: CByte/CInt/CLng/CSng/CDbl/CDec/CCur/CLngLng(Null) が Error 13 を返す（VBA では Error 94）
{
    const assertNull94 = (expr: string) => {
        let errNum = 0;
        try { evalExpr(expr); } catch (e: any) { errNum = e?.number ?? e?.errorCode ?? 0; }
        assert.strictEqual(errNum, 94, `${expr} should throw Error 94`);
    };
    assertNull94('CByte(Null)');
    assertNull94('CInt(Null)');
    assertNull94('CLng(Null)');
    assertNull94('CSng(Null)');
    assertNull94('CDbl(Null)');
    assertNull94('CDec(Null)');
    assertNull94('CCur(Null)');
    console.log('[PASS] Bug BD: CByte/CInt/CLng/CSng/CDbl/CDec/CCur(Null) = Error 94');
}

console.log('\n✅ Built-in Functions: 全テスト通過');
