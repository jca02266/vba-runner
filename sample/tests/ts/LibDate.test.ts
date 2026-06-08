import { evalVBASingle, assert, vbaTrue, vbaFalse } from '../../../test-libs/test-runner';
import * as fs from 'fs';
import * as path from 'path';

const libCode = fs.readFileSync(
    path.join(import.meta.dirname, '../../src/vba/LibDate.bas'),
    'utf-8'
);

// LibDate.bas + DateSerial ヘルパー付き評価器
function makeEv() {
    const code = libCode + `
    Function D(y, m, dv): D = DateSerial(y, m, dv): End Function
    `;
    return evalVBASingle(code);
}

const ev = makeEv();
function r(name: string, ...args: any[]): any {
    return ev.callProcedure(name, args);
}
// 日付シリアル値を作るショートカット
function d(y: number, m: number, day: number): number {
    return r('D', y, m, day) as number;
}

// --- IsLeapYear ---
{
    assert.strictEqual(r('IsLeapYear', 2000), vbaTrue, '2000 閏年');
    assert.strictEqual(r('IsLeapYear', 1900),  vbaFalse, '1900 非閏年（100の倍数）');
    assert.strictEqual(r('IsLeapYear', 2024), vbaTrue, '2024 閏年');
    assert.strictEqual(r('IsLeapYear', 2023),  vbaFalse, '2023 非閏年');
    console.log('[PASS] IsLeapYear');
}

// --- IsWeekend / IsWeekdayDate ---
{
    // 2024-03-15=金, 16=土, 17=日
    const fri = d(2024, 3, 15);
    const sat = d(2024, 3, 16);
    const sun = d(2024, 3, 17);

    assert.strictEqual(r('IsWeekend',    fri),  vbaFalse, '金曜は週末でない');
    assert.strictEqual(r('IsWeekend',    sat), vbaTrue, '土曜は週末');
    assert.strictEqual(r('IsWeekend',    sun), vbaTrue, '日曜は週末');
    assert.strictEqual(r('IsWeekdayDate', fri), vbaTrue, '金曜は平日');
    assert.strictEqual(r('IsWeekdayDate', sat),  vbaFalse, '土曜は平日でない');
    console.log('[PASS] IsWeekend / IsWeekdayDate');
}

// --- DaysInMonth ---
{
    assert.strictEqual(r('DaysInMonth', 2024, 2), 29, '2024/02 = 29日（閏年）');
    assert.strictEqual(r('DaysInMonth', 2023, 2), 28, '2023/02 = 28日');
    assert.strictEqual(r('DaysInMonth', 2024, 1), 31, '1月 = 31日');
    assert.strictEqual(r('DaysInMonth', 2024, 4), 30, '4月 = 30日');
    console.log('[PASS] DaysInMonth');
}

// --- StartOfMonth / EndOfMonth ---
{
    const mar15 = d(2024, 3, 15);
    const som = r('StartOfMonth', mar15);
    assert.strictEqual(r('Year',  som), 2024, 'SOM year=2024');
    assert.strictEqual(r('Month', som), 3,    'SOM month=3');
    assert.strictEqual(r('Day',   som), 1,    'SOM day=1');

    const eom = r('EndOfMonth', mar15);
    assert.strictEqual(r('Day',  eom), 31, 'EOM Mar=31');
    const eomFeb = r('EndOfMonth', d(2024, 2, 1));
    assert.strictEqual(r('Day', eomFeb), 29, 'EOM Feb2024=29');
    console.log('[PASS] StartOfMonth / EndOfMonth');
}

// --- QuarterOf / StartOfQuarter / EndOfQuarter ---
{
    assert.strictEqual(r('QuarterOf', d(2024,  1, 15)), 1, 'Jan=Q1');
    assert.strictEqual(r('QuarterOf', d(2024,  4,  1)), 2, 'Apr=Q2');
    assert.strictEqual(r('QuarterOf', d(2024,  7, 31)), 3, 'Jul=Q3');
    assert.strictEqual(r('QuarterOf', d(2024, 10,  1)), 4, 'Oct=Q4');

    const q1start = r('StartOfQuarter', d(2024, 2, 15));
    assert.strictEqual(r('Month', q1start), 1, 'Q1 start month=1');
    assert.strictEqual(r('Day',   q1start), 1, 'Q1 start day=1');

    const q1end = r('EndOfQuarter', d(2024, 2, 15));
    assert.strictEqual(r('Month', q1end),  3, 'Q1 end month=3');
    assert.strictEqual(r('Day',   q1end), 31, 'Q1 end day=31');
    console.log('[PASS] QuarterOf / StartOfQuarter / EndOfQuarter');
}

// --- StartOfYear / EndOfYear ---
{
    const soy = r('StartOfYear', d(2024, 6, 15));
    assert.strictEqual(r('Month', soy),  1, 'SOY month=1');
    assert.strictEqual(r('Day',   soy),  1, 'SOY day=1');

    const eoy = r('EndOfYear', d(2024, 6, 15));
    assert.strictEqual(r('Month', eoy), 12, 'EOY month=12');
    assert.strictEqual(r('Day',   eoy), 31, 'EOY day=31');
    console.log('[PASS] StartOfYear / EndOfYear');
}

// --- NextWeekday / NextOrSameWeekday / PrevOrSameWeekday ---
{
    // 2024-03-15 = 金曜 (vbFriday=6)
    const fri = d(2024, 3, 15);

    // 金曜の次の月曜 = 3/18
    const nextMon = r('NextWeekday', fri, 2); // vbMonday=2
    assert.strictEqual(r('Day', nextMon), 18, '金曜の次の月曜=18日');

    // 金曜そのものを「次の金曜」として取得 → 当日を返す
    const sameFri = r('NextOrSameWeekday', fri, 6); // vbFriday=6
    assert.strictEqual(r('Day', sameFri), 15, '当日と同じ曜日はそのまま');

    // 金曜の直前の月曜 = 3/11
    const prevMon = r('PrevOrSameWeekday', fri, 2);
    assert.strictEqual(r('Day', prevMon), 11, '金曜の直前の月曜=11日');
    console.log('[PASS] NextWeekday / PrevOrSameWeekday');
}

// --- LastWeekdayOfMonth ---
{
    // 2024年3月の最終金曜 = 29日
    const lastFri = r('LastWeekdayOfMonth', 2024, 3, 6); // vbFriday=6
    assert.strictEqual(r('Day', lastFri), 29, '2024/3 最終金曜=29日');
    console.log('[PASS] LastWeekdayOfMonth');
}

// --- NthWeekdayOfMonth ---
{
    // 2024年3月の第2月曜 = 11日
    const secondMon = r('NthWeekdayOfMonth', 2024, 3, 2, 2); // 第2月曜(vbMonday=2)
    assert.strictEqual(r('Day', secondMon), 11, '2024/3 第2月曜=11日');
    console.log('[PASS] NthWeekdayOfMonth');
}

// --- WorkdaysBetween ---
{
    const mon  = d(2024, 3, 11);  // 月曜
    const fri  = d(2024, 3, 15);  // 金曜
    const nMon = d(2024, 3, 18);  // 翌月曜

    assert.strictEqual(r('WorkdaysBetween', mon, nMon),  5, '月〜翌月曜=5営業日');
    assert.strictEqual(r('WorkdaysBetween', mon, fri),   4, '月〜金(含まず)=4営業日');
    assert.strictEqual(r('WorkdaysBetween', fri, nMon),  1, '金〜翌月曜=1営業日');

    const twoWeeksLater = d(2024, 3, 25); // 2週後の月曜
    assert.strictEqual(r('WorkdaysBetween', mon, twoWeeksLater), 10, '2週間=10営業日');
    console.log('[PASS] WorkdaysBetween');
}

// --- AddWorkdays ---
{
    const mon = d(2024, 3, 11);
    const fri = d(2024, 3, 15);

    // 月曜+5営業日=翌月曜(18日)
    assert.strictEqual(r('Day', r('AddWorkdays', mon, 5)), 18, '月曜+5営業日=18日');
    // 金曜+1営業日=翌月曜(18日)
    assert.strictEqual(r('Day', r('AddWorkdays', fri, 1)), 18, '金曜+1営業日=18日');
    // 月曜-1営業日=前金曜(8日)
    assert.strictEqual(r('Day', r('AddWorkdays', mon, -1)), 8, '月曜-1営業日=8日');
    console.log('[PASS] AddWorkdays');
}

// --- ISOWeekNumber ---
{
    assert.strictEqual(r('ISOWeekNumber', d(2024,  1,  1)),  1, '2024-01-01=W1');
    assert.strictEqual(r('ISOWeekNumber', d(2024,  3, 15)), 11, '2024-03-15=W11');
    assert.strictEqual(r('ISOWeekNumber', d(2024, 12, 30)),  1, '2024-12-30=W1(2025年第1週)');
    console.log('[PASS] ISOWeekNumber');
}

// --- ToDateString / ParseDateString / ToDateKey ---
{
    const mar15 = d(2024, 3, 15);
    assert.strictEqual(r('ToDateString',  mar15), '2024-03-15', 'ToDateString');
    assert.strictEqual(r('ToDateKey',     mar15), '20240315',   'ToDateKey');

    const parsed = r('ParseDateString', '2024-03-15');
    assert.strictEqual(r('Year',  parsed), 2024, 'ParseDateString year');
    assert.strictEqual(r('Month', parsed),    3, 'ParseDateString month');
    assert.strictEqual(r('Day',   parsed),   15, 'ParseDateString day');
    console.log('[PASS] ToDateString / ParseDateString / ToDateKey');
}

// --- IsBetweenDates ---
{
    const d1  = d(2024,  1,  1);
    const d2  = d(2024, 12, 31);
    const mid = d(2024,  6, 15);

    assert.strictEqual(r('IsBetweenDates', mid,              d1, d2), vbaTrue, '期間内');
    assert.strictEqual(r('IsBetweenDates', d1,               d1, d2), vbaTrue, '開始日は含む');
    assert.strictEqual(r('IsBetweenDates', d2,               d1, d2), vbaTrue, '終了日は含む');
    assert.strictEqual(r('IsBetweenDates', d(2023, 12, 31), d1, d2),  vbaFalse, '期間前は含まない');
    console.log('[PASS] IsBetweenDates');
}

// --- AgeInYears ---
{
    const birth = d(1990, 6, 15);
    assert.strictEqual(r('AgeInYears', birth, d(2024, 6, 14)), 33, '誕生日前日=33歳');
    assert.strictEqual(r('AgeInYears', birth, d(2024, 6, 15)), 34, '誕生日当日=34歳');
    assert.strictEqual(r('AgeInYears', birth, d(2024, 6, 16)), 34, '誕生日翌日=34歳');
    console.log('[PASS] AgeInYears');
}

console.log('\n✅ LibDate: 全テスト通過');
