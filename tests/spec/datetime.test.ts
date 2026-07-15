import { evalVBASingle, assert } from '../../test-libs/test-runner';

function evalVBA(code: string): any {
    return evalVBASingle(code);
}

function runFunc(code: string, name: string, args: any[] = []): any {
    return evalVBA(code).callProcedure(name, args);
}

// DateSerial が返す値はVBAシリアル値（Double）なので
// Year/Month/Day で逆検証する
const dsCode = `
Function TestYear(s)   : TestYear   = Year(s)   : End Function
Function TestMonth(s)  : TestMonth  = Month(s)  : End Function
Function TestDay(s)    : TestDay    = Day(s)    : End Function
Function TestHour(s)   : TestHour   = Hour(s)   : End Function
Function TestMinute(s) : TestMinute = Minute(s) : End Function
Function TestSecond(s) : TestSecond = Second(s) : End Function
Function TestWeekday(s): TestWeekday = Weekday(s): End Function
Function TestDateSerial(y, m, d)
    TestDateSerial = DateSerial(y, m, d)
End Function
Function TestTimeSerial(h, n, s)
    TestTimeSerial = TimeSerial(h, n, s)
End Function
Function TestDateAdd(iv, n, d)
    TestDateAdd = DateAdd(iv, n, d)
End Function
Function TestDateDiff(iv, d1, d2)
    TestDateDiff = DateDiff(iv, d1, d2)
End Function
Function TestDatePart(iv, d)
    TestDatePart = DatePart(iv, d)
End Function
Function TestDateValue(s)
    TestDateValue = DateValue(s)
End Function
Function TestTimeValue(s)
    TestTimeValue = TimeValue(s)
End Function
`;

// --- 1. DateSerial / Year / Month / Day ---
{
    const serial = runFunc(dsCode, 'TestDateSerial', [2024, 3, 15]);
    assert.strictEqual(runFunc(dsCode, 'TestYear',  [serial]), 2024, 'Year(DateSerial(2024,3,15))=2024');
    assert.strictEqual(runFunc(dsCode, 'TestMonth', [serial]), 3,    'Month(DateSerial(2024,3,15))=3');
    assert.strictEqual(runFunc(dsCode, 'TestDay',   [serial]), 15,   'Day(DateSerial(2024,3,15))=15');
    console.log('[PASS] DateSerial / Year / Month / Day');
}

// --- 2. 月末オーバーフロー（DateSerial の繰り上がり）---
{
    const serial = runFunc(dsCode, 'TestDateSerial', [2024, 1, 32]); // 1月32日 → 2月1日
    assert.strictEqual(runFunc(dsCode, 'TestMonth', [serial]), 2, '1月32日 → 2月');
    assert.strictEqual(runFunc(dsCode, 'TestDay',   [serial]), 1, '1月32日 → 1日');
    console.log('[PASS] DateSerial 月末オーバーフロー');
}

// --- 3. TimeSerial / Hour / Minute / Second ---
{
    const serial = runFunc(dsCode, 'TestTimeSerial', [14, 30, 45]);
    assert.strictEqual(runFunc(dsCode, 'TestHour',   [serial]), 14, 'Hour(TimeSerial(14,30,45))=14');
    assert.strictEqual(runFunc(dsCode, 'TestMinute', [serial]), 30, 'Minute(TimeSerial(14,30,45))=30');
    assert.strictEqual(runFunc(dsCode, 'TestSecond', [serial]), 45, 'Second(TimeSerial(14,30,45))=45');
    console.log('[PASS] TimeSerial / Hour / Minute / Second');
}

// --- 4. Weekday ---
{
    // 2024-03-15 は金曜日
    const serial = runFunc(dsCode, 'TestDateSerial', [2024, 3, 15]);
    assert.strictEqual(runFunc(dsCode, 'TestWeekday', [serial]), 6, 'Weekday(2024-03-15)=6(金)');
    // 2024-03-17 は日曜日
    const serial2 = runFunc(dsCode, 'TestDateSerial', [2024, 3, 17]);
    assert.strictEqual(runFunc(dsCode, 'TestWeekday', [serial2]), 1, 'Weekday(2024-03-17)=1(日)');
    console.log('[PASS] Weekday');
}

// --- 5. DateAdd ---
{
    const base = runFunc(dsCode, 'TestDateSerial', [2024, 1, 31]);
    // 月加算：1月31日 + 1ヶ月 → 2月29日（2024は閏年）
    const result = runFunc(dsCode, 'TestDateAdd', ['m', 1, base]);
    assert.strictEqual(runFunc(dsCode, 'TestMonth', [result]), 2,  'DateAdd("m",1,Jan31)→2月');
    assert.strictEqual(runFunc(dsCode, 'TestDay',   [result]), 29, 'DateAdd("m",1,Jan31)→29日（閏年）');

    // 日加算
    const base2 = runFunc(dsCode, 'TestDateSerial', [2024, 3, 15]);
    const result2 = runFunc(dsCode, 'TestDateAdd', ['d', 10, base2]);
    assert.strictEqual(runFunc(dsCode, 'TestDay', [result2]), 25, 'DateAdd("d",10,Mar15)→25日');

    // 年加算
    const result3 = runFunc(dsCode, 'TestDateAdd', ['yyyy', 1, base2]);
    assert.strictEqual(runFunc(dsCode, 'TestYear', [result3]), 2025, 'DateAdd("yyyy",1,2024)→2025');
    console.log('[PASS] DateAdd');
}

// --- 6. DateDiff ---
{
    const d1 = runFunc(dsCode, 'TestDateSerial', [2024, 1, 1]);
    const d2 = runFunc(dsCode, 'TestDateSerial', [2024, 3, 31]);

    assert.strictEqual(runFunc(dsCode, 'TestDateDiff', ['d', d1, d2]), 90, 'DateDiff("d") 1/1→3/31=90日');
    assert.strictEqual(runFunc(dsCode, 'TestDateDiff', ['m', d1, d2]), 2,  'DateDiff("m") Jan→Mar=2ヶ月');
    assert.strictEqual(runFunc(dsCode, 'TestDateDiff', ['yyyy', d1, d2]), 0, 'DateDiff("yyyy") 同年=0');

    const d3 = runFunc(dsCode, 'TestDateSerial', [2025, 1, 1]);
    assert.strictEqual(runFunc(dsCode, 'TestDateDiff', ['yyyy', d1, d3]), 1, 'DateDiff("yyyy") 1年差=1');
    assert.strictEqual(runFunc(dsCode, 'TestDateDiff', ['m', d1, d3]), 12, 'DateDiff("m") 12ヶ月差');
    console.log('[PASS] DateDiff');
}

// --- 7. DatePart ---
{
    const d = runFunc(dsCode, 'TestDateSerial', [2024, 3, 15]);
    assert.strictEqual(runFunc(dsCode, 'TestDatePart', ['yyyy', d]), 2024, 'DatePart("yyyy")');
    assert.strictEqual(runFunc(dsCode, 'TestDatePart', ['q', d]),    1,    'DatePart("q") 1Q');
    assert.strictEqual(runFunc(dsCode, 'TestDatePart', ['m', d]),    3,    'DatePart("m")');
    assert.strictEqual(runFunc(dsCode, 'TestDatePart', ['d', d]),    15,   'DatePart("d")');
    console.log('[PASS] DatePart');
}

// --- 8. DateValue ---
{
    const v = runFunc(dsCode, 'TestDateValue', ['2024-03-15']);
    assert.strictEqual(runFunc(dsCode, 'TestYear',  [v]), 2024, 'DateValue year');
    assert.strictEqual(runFunc(dsCode, 'TestMonth', [v]), 3,    'DateValue month');
    assert.strictEqual(runFunc(dsCode, 'TestDay',   [v]), 15,   'DateValue day');
    console.log('[PASS] DateValue');
}

// --- 9. TimeValue ---
{
    const v = runFunc(dsCode, 'TestTimeValue', ['14:30:45']);
    assert.strictEqual(runFunc(dsCode, 'TestHour',   [v]), 14, 'TimeValue hour');
    assert.strictEqual(runFunc(dsCode, 'TestMinute', [v]), 30, 'TimeValue minute');
    assert.strictEqual(runFunc(dsCode, 'TestSecond', [v]), 45, 'TimeValue second');
    console.log('[PASS] TimeValue');
}

// --- 10. Now / Date / Timer（値の型と範囲だけ確認）---
{
    const code2 = `
    Function TestNow()   : TestNow   = Now()   : End Function
    Function TestDate()  : TestDate  = Date()  : End Function
    Function TestTimer() : TestTimer = Timer() : End Function
    `;
    const nowVal   = runFunc(code2, 'TestNow');
    const dateVal  = runFunc(code2, 'TestDate');
    const timerVal = runFunc(code2, 'TestTimer');
    assert.strictEqual(nowVal instanceof Object && nowVal.__isVbaDate__, true, 'Now() returns VbaDate');
    assert.strictEqual(dateVal instanceof Object && dateVal.__isVbaDate__, true, 'Date() returns VbaDate');
    assert.strictEqual(typeof timerVal, 'number', 'Timer() returns number');
    assert.strictEqual(timerVal >= 0 && timerVal < 86400, true, 'Timer() 0..86400');
    console.log('[PASS] Now / Date / Timer');
}

// --- Bug #25-6: DatePart 第3引数 firstdayofweek ---
{
    const runDP = (expr: string) => {
        const ev2 = evalVBASingle(`Function T(): T = ${expr}: End Function`);
        return ev2.callProcedure('T', []);
    };
    // 2024/01/01 は月曜日
    // 'w': 日曜始まり(fdow=1)なら月曜=2、月曜始まり(fdow=2)なら月曜=1
    assert.strictEqual(runDP('DatePart("w", #2024/01/01#, 1)'), 2, 'DatePart w 日曜始まり: 月曜=2');
    assert.strictEqual(runDP('DatePart("w", #2024/01/01#, 2)'), 1, 'DatePart w 月曜始まり: 月曜=1');
    assert.strictEqual(runDP('DatePart("w", #2024/01/07#, 1)'), 1, 'DatePart w 日曜始まり: 日曜=1');
    assert.strictEqual(runDP('DatePart("w", #2024/01/07#, 2)'), 7, 'DatePart w 月曜始まり: 日曜=7');
    // 'ww': 週番号も firstdayofweek に従う
    const ww_sun = runDP('DatePart("ww", #2024/01/01#, 1)') as number;
    const ww_mon = runDP('DatePart("ww", #2024/01/01#, 2)') as number;
    assert.strictEqual(ww_mon, 1, 'DatePart ww 月曜始まり: 2024/01/01 は第1週');
    assert.ok(typeof ww_sun === 'number', 'DatePart ww 日曜始まりが数値');
    console.log('[PASS] Bug #25-6: DatePart firstdayofweek');
}

// --- Bug B: Weekday — firstdayofweek 引数 ---
{
    const runWD = (expr: string) => {
        const ev2 = evalVBASingle(`Function T(): T = ${expr}: End Function`);
        return ev2.callProcedure('T', []);
    };
    // 2026/07/15 は水曜日 (JS getDay=3)
    // vbSunday(1)基準: 日=1,月=2,火=3,水=4,木=5,金=6,土=7
    assert.strictEqual(runWD('Weekday(#2026/07/15#)'),    4, 'Weekday デフォルト(vbSunday): 水=4');
    assert.strictEqual(runWD('Weekday(#2026/07/15#, 1)'), 4, 'Weekday vbSunday(1): 水=4');
    // vbMonday(2)基準: 月=1,火=2,水=3
    assert.strictEqual(runWD('Weekday(#2026/07/15#, 2)'), 3, 'Weekday vbMonday(2): 水=3');
    // vbSaturday(7)基準: 土=1,日=2,月=3,火=4,水=5
    assert.strictEqual(runWD('Weekday(#2026/07/15#, 7)'), 5, 'Weekday vbSaturday(7): 水=5');
    // vbUseSystemDayOfWeek(0) → vbSunday として扱う
    assert.strictEqual(runWD('Weekday(#2026/07/15#, 0)'), 4, 'Weekday vbUseSystem(0): 水=4');
    console.log('[PASS] Bug B: Weekday firstdayofweek');
}

// --- Bug C: DateDiff — firstdayofweek 引数 ---
{
    const runDD = (expr: string) => {
        const ev2 = evalVBASingle(`Function T(): T = ${expr}: End Function`);
        return ev2.callProcedure('T', []);
    };
    // 非 "ww" インターバルは firstdayofweek 無関係
    assert.strictEqual(runDD('DateDiff("d", #2026/01/01#, #2026/01/31#, 2)'), 30, 'DateDiff d: firstdayofweek 無関係');
    // "ww": 2026/01/01(木) → 2026/01/31(土)
    // vbSunday(1): 01/01の週始=12/28(日), 01/31の週始=01/25(日) → 4週
    assert.strictEqual(runDD('DateDiff("ww", #2026/01/01#, #2026/01/31#, 1)'), 4, 'DateDiff ww vbSunday: 4週');
    // vbMonday(2): 01/01の週始=12/29(月), 01/31の週始=01/26(月) → 4週
    assert.strictEqual(runDD('DateDiff("ww", #2026/01/01#, #2026/01/31#, 2)'), 4, 'DateDiff ww vbMonday: 4週');
    // デフォルト(1引数省略)も正常動作
    assert.strictEqual(runDD('DateDiff("ww", #2026/01/01#, #2026/01/31#)'), 4, 'DateDiff ww デフォルト: 4週');
    console.log('[PASS] Bug C: DateDiff firstdayofweek');
}

console.log('\n✅ DateTime Module: 全テスト通過');
