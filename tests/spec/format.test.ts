import { evalVBASingle, assert } from '../../test-libs/test-runner';

const code = `
    Function TestFormatDate()
        TestFormatDate = Format(#2025/05/10#, "yyyy-mm-dd")
    End Function

    Function TestFormatNumber()
        TestFormatNumber = Format(1234.567, "0.00")
    End Function

    Function TestFormatCurrency()
        TestFormatCurrency = Format(1234.567, "#,##0.00")
    End Function
`;

const ev = evalVBASingle(code);

console.log('[Test Suite] Format の検証');

assert.strictEqual(ev.callProcedure('TestFormatDate', []), '2025-05-10', 'Format(Date, "yyyy-mm-dd")');
assert.strictEqual(ev.callProcedure('TestFormatNumber', []), '1234.57', 'Format(Number, "0.00")');
const currency = ev.callProcedure('TestFormatCurrency', []);
assert.ok(currency.includes('1,234.57') || currency.includes('1234.57'), 'Format(Number, "#,##0.00")');

// --- 時刻フォーマットのバグ修正確認 ---
const timeCode = `
    Function TestHH()
        TestHH = Format(#2023-06-15 09:30:45#, "HH:mm:ss")
    End Function
    Function TestHhNn()
        TestHhNn = Format(#2023-06-15 09:30:45#, "hh:nn:ss")
    End Function
    Function TestIsoFormat()
        TestIsoFormat = Format(#2023-01-15 12:34:56#, "yyyy-mm-ddTHH:mm:ss")
    End Function
    Function TestMmIsMonth()
        TestMmIsMonth = Format(#2023-06-15#, "yyyy-mm-dd")
    End Function
    Function TestMmIsMinute()
        TestMmIsMinute = Format(#2023-06-15 09:07:05#, "hh:mm:ss")
    End Function
`;
const evTime = evalVBASingle(timeCode);

assert.strictEqual(evTime.callProcedure('TestHH', []), '09:30:45',
    'Format HH:mm:ss — HH(大文字)が時刻、mm が文脈上「分」になる');
assert.strictEqual(evTime.callProcedure('TestHhNn', []), '09:30:45',
    'Format hh:nn:ss — nn は常に分');
assert.strictEqual(evTime.callProcedure('TestIsoFormat', []), '2023-01-15T12:34:56',
    'Format "yyyy-mm-ddTHH:mm:ss" — 最初の mm は月、HH 後の mm は分');
assert.strictEqual(evTime.callProcedure('TestMmIsMonth', []), '2023-06-15',
    'Format yyyy-mm-dd — mm は月');
assert.strictEqual(evTime.callProcedure('TestMmIsMinute', []), '09:07:05',
    'Format hh:mm:ss — mm は分（ゼロ埋め）');

console.log('[PASS] Format 時刻: HH大文字・文脈依存mm・UTC整合性');

// --- 数値フォーマット: 零埋め・# プレースホルダー ---
const numCode = `
    Function TestZeroPad()
        TestZeroPad = Format(42, "000")
    End Function
    Function TestZeroPadZero()
        TestZeroPadZero = Format(0, "000")
    End Function
    Function TestHashZero()
        TestHashZero = Format(0, "#")
    End Function
    Function TestHashDecimalZero()
        TestHashDecimalZero = Format(0, "#.##")
    End Function
    Function TestHashLeadingZero()
        TestHashLeadingZero = Format(0.5, "#.#")
    End Function
    Function TestHashWithValue()
        TestHashWithValue = Format(42, "#")
    End Function
    Function TestNegativeZeroPad()
        TestNegativeZeroPad = Format(-42, "000")
    End Function
    Function TestPercent()
        TestPercent = Format(0.155, "0.00%")
    End Function
    Function TestThousands()
        TestThousands = Format(1234567, "#,##0")
    End Function
`;
const evNum = evalVBASingle(numCode);

assert.strictEqual(evNum.callProcedure('TestZeroPad', []), '042', 'Format(42, "000") → "042"');
assert.strictEqual(evNum.callProcedure('TestZeroPadZero', []), '000', 'Format(0, "000") → "000"');
assert.strictEqual(evNum.callProcedure('TestHashZero', []), '', 'Format(0, "#") → "" (# suppresses zero)');
assert.strictEqual(evNum.callProcedure('TestHashDecimalZero', []), '', 'Format(0, "#.##") → "" (# suppresses zero)');
assert.strictEqual(evNum.callProcedure('TestHashLeadingZero', []), '.5', 'Format(0.5, "#.#") → ".5" (# suppresses leading zero)');
assert.strictEqual(evNum.callProcedure('TestHashWithValue', []), '42', 'Format(42, "#") → "42"');
assert.strictEqual(evNum.callProcedure('TestNegativeZeroPad', []), '-042', 'Format(-42, "000") → "-042"');
assert.strictEqual(evNum.callProcedure('TestPercent', []), '15.50%', 'Format(0.155, "0.00%") → "15.50%"');
assert.strictEqual(evNum.callProcedure('TestThousands', []), '1,234,567', 'Format(1234567, "#,##0") → "1,234,567"');

console.log('[PASS] Format 数値: 零埋め・# プレースホルダー・パーセント・千の位');

// --- 数値フォーマット: リテラル文字・複数セクション・科学記数法 ---
const numCode2 = `
    Function TestLiteralPrefix()
        TestLiteralPrefix = Format(42, "$#,##0.00")
    End Function
    Function TestLiteralSuffix()
        TestLiteralSuffix = Format(42, "#,##0 円")
    End Function
    Function TestSectionPos()
        TestSectionPos = Format(5, "#;(#);Zero")
    End Function
    Function TestSectionNeg()
        TestSectionNeg = Format(-5, "#;(#);Zero")
    End Function
    Function TestSectionZero()
        TestSectionZero = Format(0, "#;(#);Zero")
    End Function
    Function TestSectionNeg2()
        TestSectionNeg2 = Format(-5, "#;(#)")
    End Function
    Function TestSciLarge()
        TestSciLarge = Format(12345.6, "0.00E+00")
    End Function
    Function TestSciSmall()
        TestSciSmall = Format(0.00123, "0.00E+00")
    End Function
    Function TestSciZero()
        TestSciZero = Format(0, "0.00E+00")
    End Function
    Function TestEscape()
        TestEscape = Format(0, "$#,##0;;\\Z\\e\\r\\o")
    End Function
`;
const evNum2 = evalVBASingle(numCode2);

assert.strictEqual(evNum2.callProcedure('TestLiteralPrefix', []), '$42.00',   'Format(42, "$#,##0.00") → "$42.00"');
assert.strictEqual(evNum2.callProcedure('TestLiteralSuffix', []), '42 円',    'Format(42, "#,##0 円") → "42 円"');
assert.strictEqual(evNum2.callProcedure('TestSectionPos',    []), '5',        'Format(5, "#;(#);Zero") → "5"');
assert.strictEqual(evNum2.callProcedure('TestSectionNeg',    []), '(5)',      'Format(-5, "#;(#);Zero") → "(5)"');
assert.strictEqual(evNum2.callProcedure('TestSectionZero',   []), 'Zero',     'Format(0, "#;(#);Zero") → "Zero"');
assert.strictEqual(evNum2.callProcedure('TestSectionNeg2',   []), '(5)',      'Format(-5, "#;(#)") → "(5)"');
assert.strictEqual(evNum2.callProcedure('TestSciLarge',      []), '1.23E+04', 'Format(12345.6, "0.00E+00") → "1.23E+04"');
assert.strictEqual(evNum2.callProcedure('TestSciSmall',      []), '1.23E-03', 'Format(0.00123, "0.00E+00") → "1.23E-03"');
assert.strictEqual(evNum2.callProcedure('TestSciZero',       []), '0.00E+00', 'Format(0, "0.00E+00") → "0.00E+00"');
assert.strictEqual(evNum2.callProcedure('TestEscape',        []), 'Zero',     'Format(0, "$#,##0;;\\Z\\e\\r\\o") → "Zero"');

console.log('[PASS] Format 数値: リテラル・複数セクション・科学記数法・\\エスケープ');

// --- 文字列フォーマット ---
const strCode = `
    Function TestUpper()
        TestUpper = Format("hello", ">")
    End Function
    Function TestLower()
        TestLower = Format("HELLO", "<")
    End Function
    Function TestAtRight()
        TestAtRight = Format("hi", "@@@@@")
    End Function
    Function TestAtLeft()
        TestAtLeft = Format("hi", "!@@@@@")
    End Function
    Function TestAmpersand()
        TestAmpersand = Format("hi", "&&&&&")
    End Function
`;
const evStr = evalVBASingle(strCode);

assert.strictEqual(evStr.callProcedure('TestUpper',     []), 'HELLO', 'Format("hello", ">") → "HELLO"');
assert.strictEqual(evStr.callProcedure('TestLower',     []), 'hello', 'Format("HELLO", "<") → "hello"');
assert.strictEqual(evStr.callProcedure('TestAtRight',   []), '   hi', 'Format("hi", "@@@@@") → "   hi"');
assert.strictEqual(evStr.callProcedure('TestAtLeft',    []), 'hi   ', 'Format("hi", "!@@@@@") → "hi   "');
assert.strictEqual(evStr.callProcedure('TestAmpersand', []), 'hi',    'Format("hi", "&&&&&") → "hi"');

console.log('[PASS] Format 文字列: >, <, @, &, !');

// --- 日付フォーマット: 新トークン ---
const dateCode2 = String.raw`
    Function TestQ()
        TestQ = Format(#2025/05/10#, "q")
    End Function
    Function TestW()
        TestW = Format(#2025/01/01#, "w")
    End Function
    Function TestY()
        TestY = Format(#2025/01/31#, "y")
    End Function
    Function TestDdddd()
        TestDdddd = Format(#2025/05/10#, "ddddd")
    End Function
    Function TestTtttt()
        TestTtttt = Format(#2025/05/10 09:30:45#, "ttttt")
    End Function
    Function TestEscapeDate()
        TestEscapeDate = Format(#2025/05/10#, "\Y""Year""yyyy")
    End Function
`;
const evDate2 = evalVBASingle(dateCode2);

assert.strictEqual(evDate2.callProcedure('TestQ',          []), '2',         'Format(#2025/05/10#, "q") → "2"');
assert.strictEqual(evDate2.callProcedure('TestW',          []), '4',         'Format(#2025/01/01#, "w") → "4" (水曜=4)');
assert.strictEqual(evDate2.callProcedure('TestY',          []), '31',        'Format(#2025/01/31#, "y") → "31"');
assert.strictEqual(evDate2.callProcedure('TestDdddd',      []), '5/10/25',   'Format(#2025/05/10#, "ddddd") → "5/10/25"');
assert.strictEqual(evDate2.callProcedure('TestTtttt',      []), '9:30:45',   'Format(#... 09:30:45#, "ttttt") → "9:30:45"');
assert.strictEqual(evDate2.callProcedure('TestEscapeDate', []), 'YYear2025', 'Format(date, "\\Y\\"Year\\"yyyy") → "YYear2025"');

console.log('[PASS] Format 日付: q, w, y, ddddd, ttttt, \\エスケープ・"text"リテラル');

// --- Bug BG: Format() が Math.round を使い、銀行家丸め（VBA 仕様）ではなかった ---
const roundCode = `
    Function TestHalfEven() As String
        TestHalfEven = Format(2.5, "0") & "," & Format(3.5, "0") & "," & Format(1234.5, "0") & "," & Format(0.5, "0")
    End Function
    Function TestHalfEvenDec() As String
        TestHalfEvenDec = Format(2.25, "0.0") & "," & Format(2.35, "0.0")
    End Function
`;
const evRound = evalVBASingle(roundCode);
assert.strictEqual(evRound.callProcedure('TestHalfEven', []), '2,4,1234,0', 'Format 銀行家丸め: 2.5→2, 3.5→4, 1234.5→1234, 0.5→0');
assert.strictEqual(evRound.callProcedure('TestHalfEvenDec', []), '2.2,2.4', 'Format 銀行家丸め 小数: 2.25→2.2, 2.35→2.4');
console.log('[PASS] Bug BG: Format() 銀行家丸め');

// --- Bug BQ: Format(time, "Long Time") が 24 時間表記になっていた、
//             Format(date, "Short Date") が ISO 形式 (yyyy/MM/dd) になっていた ---
{
    const ev = evalVBASingle('');
    assert.strictEqual(ev.evalExpression('Format(TimeSerial(10, 30, 45), "Long Time")'), '10:30:45 AM', 'Long Time 午前');
    assert.strictEqual(ev.evalExpression('Format(TimeSerial(22, 30, 0), "Long Time")'), '10:30:00 PM', 'Long Time 午後');
    assert.strictEqual(ev.evalExpression('Format(TimeSerial(12, 0, 0), "Long Time")'), '12:00:00 PM', 'Long Time 正午');
    assert.strictEqual(ev.evalExpression('Format(TimeSerial(0, 0, 0), "Long Time")'), '12:00:00 AM', 'Long Time 深夜0時');
    assert.strictEqual(ev.evalExpression('Format(DateSerial(2024, 1, 5), "Short Date")'), '1/5/2024', 'Short Date: M/D/YYYY');
    assert.strictEqual(ev.evalExpression('Format(DateSerial(2024, 12, 31), "Short Date")'), '12/31/2024', 'Short Date: 12桁月');
    console.log('[PASS] Bug BQ: Long Time AM/PM と Short Date M/D/YYYY');
}

console.log('\n✅ Format: 全テスト通過');
