import { Lexer } from '../../src/engine/lexer';
import { Parser } from '../../src/engine/parser';
import { Evaluator } from '../../src/engine/evaluator';
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

// Test 1: Jan 31 + 1 month = Feb 28 (non-leap year)
{
    const code = `
    Function Test1()
        Dim d As Date
        d = DateAdd("m", 1, #1/31/2023#)
        Test1 = Day(d) & "/" & Month(d) & "/" & Year(d)
    End Function
    `;
    const result = runFunc(code, 'Test1');
    assert.strictEqual(result, '28/2/2023', 'Jan 31 + 1 month should be Feb 28 in non-leap year');
    console.log('[PASS] Test 1: Jan 31 + 1 month = Feb 28 (non-leap year)');
}

// Test 2: Jan 31 + 1 month = Feb 29 (leap year)
{
    const code = `
    Function Test2()
        Dim d As Date
        d = DateAdd("m", 1, #1/31/2024#)
        Test2 = Day(d) & "/" & Month(d) & "/" & Year(d)
    End Function
    `;
    const result = runFunc(code, 'Test2');
    assert.strictEqual(result, '29/2/2024', 'Jan 31 + 1 month should be Feb 29 in leap year');
    console.log('[PASS] Test 2: Jan 31 + 1 month = Feb 29 (leap year)');
}

// Test 3: Jan 31 + 2 months = Mar 31 (month with 31 days)
{
    const code = `
    Function Test3()
        Dim d As Date
        d = DateAdd("m", 2, #1/31/2023#)
        Test3 = Day(d) & "/" & Month(d) & "/" & Year(d)
    End Function
    `;
    const result = runFunc(code, 'Test3');
    assert.strictEqual(result, '31/3/2023', 'Jan 31 + 2 months should be Mar 31');
    console.log('[PASS] Test 3: Jan 31 + 2 months = Mar 31');
}

// Test 4: Mar 31 + 1 month = Apr 30 (April has 30 days)
{
    const code = `
    Function Test4()
        Dim d As Date
        d = DateAdd("m", 1, #3/31/2023#)
        Test4 = Day(d) & "/" & Month(d) & "/" & Year(d)
    End Function
    `;
    const result = runFunc(code, 'Test4');
    assert.strictEqual(result, '30/4/2023', 'Mar 31 + 1 month should be Apr 30');
    console.log('[PASS] Test 4: Mar 31 + 1 month = Apr 30 (April has 30 days)');
}

// Test 5: May 31 + 1 month = Jun 30 (June has 30 days)
{
    const code = `
    Function Test5()
        Dim d As Date
        d = DateAdd("m", 1, #5/31/2023#)
        Test5 = Day(d) & "/" & Month(d) & "/" & Year(d)
    End Function
    `;
    const result = runFunc(code, 'Test5');
    assert.strictEqual(result, '30/6/2023', 'May 31 + 1 month should be Jun 30');
    console.log('[PASS] Test 5: May 31 + 1 month = Jun 30 (June has 30 days)');
}

// Test 6: Aug 31 + 1 month = Sep 30 (September has 30 days)
{
    const code = `
    Function Test6()
        Dim d As Date
        d = DateAdd("m", 1, #8/31/2023#)
        Test6 = Day(d) & "/" & Month(d) & "/" & Year(d)
    End Function
    `;
    const result = runFunc(code, 'Test6');
    assert.strictEqual(result, '30/9/2023', 'Aug 31 + 1 month should be Sep 30');
    console.log('[PASS] Test 6: Aug 31 + 1 month = Sep 30 (September has 30 days)');
}

// Test 7: Oct 31 + 1 month = Nov 30 (November has 30 days)
{
    const code = `
    Function Test7()
        Dim d As Date
        d = DateAdd("m", 1, #10/31/2023#)
        Test7 = Day(d) & "/" & Month(d) & "/" & Year(d)
    End Function
    `;
    const result = runFunc(code, 'Test7');
    assert.strictEqual(result, '30/11/2023', 'Oct 31 + 1 month should be Nov 30');
    console.log('[PASS] Test 7: Oct 31 + 1 month = Nov 30 (November has 30 days)');
}

// Test 8: Dec 31 + 1 month (wraps to next year) = Jan 31
{
    const code = `
    Function Test8()
        Dim d As Date
        d = DateAdd("m", 1, #12/31/2023#)
        Test8 = Day(d) & "/" & Month(d) & "/" & Year(d)
    End Function
    `;
    const result = runFunc(code, 'Test8');
    assert.strictEqual(result, '31/1/2024', 'Dec 31 + 1 month should be Jan 31 (next year)');
    console.log('[PASS] Test 8: Dec 31 + 1 month = Jan 31 (next year)');
}

// Test 9: Jan 31 - 1 month = Dec 31 (previous year)
{
    const code = `
    Function Test9()
        Dim d As Date
        d = DateAdd("m", -1, #1/31/2024#)
        Test9 = Day(d) & "/" & Month(d) & "/" & Year(d)
    End Function
    `;
    const result = runFunc(code, 'Test9');
    assert.strictEqual(result, '31/12/2023', 'Jan 31 - 1 month should be Dec 31 (previous year)');
    console.log('[PASS] Test 9: Jan 31 - 1 month = Dec 31 (previous year)');
}

// Test 10: Jan 30 + 1 month = Feb 29 (leap year, day capped to last day)
{
    const code = `
    Function Test10()
        Dim d As Date
        d = DateAdd("m", 1, #1/30/2024#)
        Test10 = Day(d) & "/" & Month(d) & "/" & Year(d)
    End Function
    `;
    const result = runFunc(code, 'Test10');
    assert.strictEqual(result, '29/2/2024', 'Jan 30 + 1 month should be Feb 29 in leap year (day capped to 29)');
    console.log('[PASS] Test 10: Jan 30 + 1 month = Feb 29 (leap year, day capped)');
}

// Test 11: Jan 29 + 1 month = Feb 29 (leap year, exact match)
{
    const code = `
    Function Test11()
        Dim d As Date
        d = DateAdd("m", 1, #1/29/2024#)
        Test11 = Day(d) & "/" & Month(d) & "/" & Year(d)
    End Function
    `;
    const result = runFunc(code, 'Test11');
    assert.strictEqual(result, '29/2/2024', 'Jan 29 + 1 month should be Feb 29 in leap year');
    console.log('[PASS] Test 11: Jan 29 + 1 month = Feb 29 (leap year)');
}

// Test 12: Multiple month additions from end-of-month date
{
    const code = `
    Function Test12()
        Dim d As Date
        d = DateAdd("m", 3, #1/31/2023#)
        Test12 = Day(d) & "/" & Month(d) & "/" & Year(d)
    End Function
    `;
    const result = runFunc(code, 'Test12');
    assert.strictEqual(result, '30/4/2023', 'Jan 31 + 3 months should be Apr 30');
    console.log('[PASS] Test 12: Jan 31 + 3 months = Apr 30');
}

// Test 13: Year-end to month-end across year boundary
{
    const code = `
    Function Test13()
        Dim d As Date
        d = DateAdd("m", 13, #1/31/2023#)
        Test13 = Day(d) & "/" & Month(d) & "/" & Year(d)
    End Function
    `;
    const result = runFunc(code, 'Test13');
    assert.strictEqual(result, '29/2/2024', 'Jan 31 + 13 months should be Feb 29, 2024 (leap year)');
    console.log('[PASS] Test 13: Jan 31 + 13 months = Feb 29, 2024 (leap year)');
}

// Test 14: Negative month offset from month-end
{
    const code = `
    Function Test14()
        Dim d As Date
        d = DateAdd("m", -2, #5/31/2023#)
        Test14 = Day(d) & "/" & Month(d) & "/" & Year(d)
    End Function
    `;
    const result = runFunc(code, 'Test14');
    assert.strictEqual(result, '31/3/2023', 'May 31 - 2 months should be Mar 31');
    console.log('[PASS] Test 14: May 31 - 2 months = Mar 31');
}

// Test 15: Verify month-end rollover from May to June
{
    const code = `
    Function Test15()
        Dim d As Date
        d = DateAdd("m", 2, #3/31/2023#)
        Test15 = Day(d) & "/" & Month(d) & "/" & Year(d)
    End Function
    `;
    const result = runFunc(code, 'Test15');
    assert.strictEqual(result, '31/5/2023', 'Mar 31 + 2 months should be May 31');
    console.log('[PASS] Test 15: Mar 31 + 2 months = May 31');
}

console.log('\n✅ DateAdd Month-End Rollover: 全テスト通過');
