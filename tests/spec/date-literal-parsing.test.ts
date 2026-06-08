import { Evaluator } from '../../src/engine/evaluator';
import { evalVBASingle, assert } from '../../test-libs/test-runner';

function evalVBA(code: string): Evaluator {
    return evalVBASingle(code);
}

function runFunc(code: string, name: string, args: any[] = []): any {
    return evalVBA(code).callProcedure(name, args);
}

// Helper: Get month from date value
function getMonth(dateValue: any): number {
    const d = new Date(dateValue * 86400 * 1000 + new Date('1899-12-30').getTime());
    return d.getMonth() + 1;
}

function getDay(dateValue: any): number {
    const d = new Date(dateValue * 86400 * 1000 + new Date('1899-12-30').getTime());
    return d.getDate();
}

function getYear(dateValue: any): number {
    const d = new Date(dateValue * 86400 * 1000 + new Date('1899-12-30').getTime());
    return d.getFullYear();
}

// Test 1: Standard mm/dd/yyyy interpretation (#1/2/2024#)
{
    const code = `
        Function Test1() As Variant
            Test1 = #1/2/2024#
        End Function
    `;

    try {
        const result = runFunc(code, 'Test1');
        const month = getMonth(result);
        const day = getDay(result);
        const year = getYear(result);
        assert.strictEqual(month, 1, 'Date #1/2/2024# has month 1');
        assert.strictEqual(day, 2, 'Date #1/2/2024# has day 2');
        assert.strictEqual(year, 2024, 'Date #1/2/2024# has year 2024');
        console.log('[PASS] Standard mm/dd/yyyy interpretation');
    } catch (e: any) {
        console.log('[FAIL] Standard mm/dd/yyyy interpretation:', e.message);
    }
}

// Test 2: Ambiguous date resolves by month validity (#13/2/2024# => February 13)
{
    const code = `
        Function Test2() As Variant
            Test2 = #13/2/2024#
        End Function
    `;

    try {
        const result = runFunc(code, 'Test2');
        const month = getMonth(result);
        const day = getDay(result);
        // 13 cannot be a month, so second rule applies: M=2 is month, L=13 is day
        assert.strictEqual(month, 2, 'Date #13/2/2024# resolves to February');
        assert.strictEqual(day, 13, 'Date #13/2/2024# has day 13');
        console.log('[PASS] Ambiguous date resolves by month validity');
    } catch (e: any) {
        console.log('[FAIL] Ambiguous date resolves by month validity:', e.message);
    }
}

// Test 3: Date with two-digit year (#1/2/24#)
{
    const code = `
        Function Test3() As Variant
            Test3 = #1/2/24#
        End Function
    `;

    try {
        const result = runFunc(code, 'Test3');
        const year = getYear(result);
        // Two-digit year should be interpreted as 2024 (current year sliding window)
        assert.strictEqual(year >= 2000, true, 'Date #1/2/24# has valid year');
        console.log('[PASS] Date with two-digit year');
    } catch (e: any) {
        console.log('[FAIL] Date with two-digit year:', e.message);
    }
}

// Test 4: Date with English month name (#January 2, 2024#)
{
    const code = `
        Function Test4() As Variant
            Test4 = #January 2, 2024#
        End Function
    `;

    try {
        const result = runFunc(code, 'Test4');
        const month = getMonth(result);
        const day = getDay(result);
        const year = getYear(result);
        assert.strictEqual(month, 1, 'Date with month name has month 1');
        assert.strictEqual(day, 2, 'Date with month name has day 2');
        assert.strictEqual(year, 2024, 'Date with month name has year 2024');
        console.log('[PASS] Date with English month name');
    } catch (e: any) {
        console.log('[FAIL] Date with English month name:', e.message);
    }
}

// Test 5: Date with abbreviated month name (#Jan 2, 2024#)
{
    const code = `
        Function Test5() As Variant
            Test5 = #Jan 2, 2024#
        End Function
    `;

    try {
        const result = runFunc(code, 'Test5');
        const month = getMonth(result);
        const day = getDay(result);
        assert.strictEqual(month, 1, 'Date with abbreviated month has month 1');
        assert.strictEqual(day, 2, 'Date with abbreviated month has day 2');
        console.log('[PASS] Date with abbreviated month name');
    } catch (e: any) {
        console.log('[FAIL] Date with abbreviated month name:', e.message);
    }
}

// Test 6: Date with dashes (#1-2-2024#)
{
    const code = `
        Function Test6() As Variant
            Test6 = #1-2-2024#
        End Function
    `;

    try {
        const result = runFunc(code, 'Test6');
        const month = getMonth(result);
        const day = getDay(result);
        assert.strictEqual(month, 1, 'Date with dashes has correct month');
        assert.strictEqual(day, 2, 'Date with dashes has correct day');
        console.log('[PASS] Date with dashes as separator');
    } catch (e: any) {
        console.log('[FAIL] Date with dashes as separator:', e.message);
    }
}

// Test 7: Date operations preserve locale-independence
{
    const code = `
        Function Test7() As Boolean
            Dim d1 As Variant, d2 As Variant
            d1 = #1/15/2024#
            d2 = #15/1/2024#
            ' d1 should be Jan 15, d2 should be Jan 15 too (day/month ambiguity resolved)
            Test7 = (d1 = d2)
        End Function
    `;

    try {
        const result = runFunc(code, 'Test7');
        // Both dates should be the same if parsing is consistent
        const vbaTrue = result === true || (result && typeof result === 'object' && result.value === -1);
        assert.strictEqual(vbaTrue, true, 'Dates parse consistently regardless of component order');
        console.log('[PASS] Date operations preserve locale-independence');
    } catch (e: any) {
        console.log('[FAIL] Date operations preserve locale-independence:', e.message);
    }
}

// Test 8: Time component in date literal (#1/2/2024 12:30:45#)
{
    const code = `
        Function Test8() As Variant
            Test8 = #1/2/2024 12:30:45#
        End Function
    `;

    try {
        const result = runFunc(code, 'Test8');
        const month = getMonth(result);
        const day = getDay(result);
        const year = getYear(result);
        assert.strictEqual(month, 1, 'Date with time has correct month');
        assert.strictEqual(day, 2, 'Date with time has correct day');
        assert.strictEqual(year, 2024, 'Date with time has correct year');
        console.log('[PASS] Time component in date literal');
    } catch (e: any) {
        console.log('[FAIL] Time component in date literal:', e.message);
    }
}

// Test 9: Date comparison (#1/2/2024# vs #1/2/2024#)
{
    const code = `
        Function Test9() As Boolean
            Test9 = (#1/2/2024# = #1/2/2024#)
        End Function
    `;

    try {
        const result = runFunc(code, 'Test9');
        const vbaTrue = result === true || (result && typeof result === 'object' && result.value === -1);
        assert.strictEqual(vbaTrue, true, 'Identical date literals compare equal');
        console.log('[PASS] Date comparison');
    } catch (e: any) {
        console.log('[FAIL] Date comparison:', e.message);
    }
}

// Test 10: Real-world pattern - Date validation and arithmetic
{
    const code = `
        Function Test10() As Integer
            Dim startDate As Variant, endDate As Variant
            startDate = #1/1/2024#
            endDate = #1/31/2024#
            ' Count days between dates
            Test10 = Int(endDate - startDate)
        End Function
    `;

    try {
        const result = runFunc(code, 'Test10');
        assert.strictEqual(result, 30, 'Date arithmetic works (30 days in difference)');
        console.log('[PASS] Real-world: Date validation and arithmetic');
    } catch (e: any) {
        console.log('[FAIL] Real-world: Date validation and arithmetic:', e.message);
    }
}

console.log('\n✅ Date Literal Parsing: 全テスト完了');
