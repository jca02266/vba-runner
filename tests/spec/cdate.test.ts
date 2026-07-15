import { evalVBASingle, assert } from '../../test-libs/test-runner';
import { VbaDate } from '../../src/engine/evaluator';

const code = `
    Function TestCDateString()
        TestCDateString = CDate("2025/05/10")
    End Function

    Function TestCDateNumber()
        TestCDateNumber = CDate(45658)
    End Function

    Function TestCVDateNull()
        TestCVDateNull = CVDate(Null)
    End Function

    Function TestCVDateDate()
        TestCVDateDate = CVDate(#2025/05/10#)
    End Function
`;

const ev = evalVBASingle(code);

console.log('[Test Suite] CDate / CVDate の検証');

const dateStr = ev.callProcedure('TestCDateString', []);
assert.ok(dateStr instanceof VbaDate, 'CDate(string) should return VbaDate');
assert.strictEqual(ev.callProcedure('TypeName', [dateStr]), 'Date', 'TypeName should be Date');

const dateNum = ev.callProcedure('TestCDateNumber', []);
assert.ok(dateNum instanceof VbaDate, 'CDate(number) should return VbaDate');
assert.strictEqual(ev.callProcedure('TypeName', [dateNum]), 'Date', 'TypeName should be Date');

const vdateNull = ev.callProcedure('TestCVDateNull', []);
assert.strictEqual(ev.callProcedure('TypeName', [vdateNull]), 'Null', 'CVDate(Null) should be Null');

const vdateDate = ev.callProcedure('TestCVDateDate', []);
assert.ok(vdateDate instanceof VbaDate, 'CVDate(Date) should return VbaDate');

// Bug R: ISO "YYYY-MM-DD" 形式の文字列がタイムゾーンオフセットを含む問題
{
    // "2024-01-31" は JS では UTC midnight として解析されるため、
    // UTC+9 環境では 9:00 AM ローカル時刻 = VBA シリアル値の小数部 0.375 になっていた
    const d = ev.callProcedure('CDate', []) as any;
    const dIso = evalVBASingle(`Function F() : F = CDate("2024-01-31") : End Function`).callProcedure('F', []) as any;
    assert.ok(dIso instanceof VbaDate, 'CDate("2024-01-31") returns VbaDate');
    const frac = dIso.value - Math.floor(dIso.value);
    assert.strictEqual(frac, 0, 'CDate("2024-01-31") should have no time component (fractional=0)');
    console.log('[PASS] Bug R: ISO 日付文字列のタイムゾーンオフセット修正');
}

console.log('✅ CDate / CVDate: 全テスト通過');
