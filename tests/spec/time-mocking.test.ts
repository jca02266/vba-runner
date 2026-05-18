import { Lexer } from '../../src/engine/lexer';
import { Parser } from '../../src/engine/parser';
import { Evaluator } from '../../src/engine/evaluator';
import { VBARunner, assert } from '../../test-libs/test-runner';

function makeEv(nowFn: (() => Date) | null = null): Evaluator {
    const ev = new Evaluator(() => {});
    if (nowFn) ev.setNowFn(nowFn);
    return ev;
}

function run(ev: Evaluator, code: string, proc: string, args: any[] = []): any {
    const ast = new Parser(new Lexer(code).tokenize()).parse();
    ev.evaluate(ast);
    return ev.callProcedure(proc, args);
}

const FIXED = new Date(2024, 2, 15, 10, 30, 45); // 2024-03-15 10:30:45 local

// 1. Now() returns mocked date
{
    const ev = makeEv(() => FIXED);
    const result = run(ev, `
Function Test1() As Long
    Test1 = Year(Now())
End Function
`, 'Test1');
    assert.strictEqual(result, 2024, 'Year(Now()) with mocked date');
    console.log('[PASS] Year(Now()) =', result);
}

// 2. Date returns date-only part
{
    const ev = makeEv(() => FIXED);
    const result = run(ev, `
Function Test2() As Long
    Test2 = Month(Date())
End Function
`, 'Test2');
    assert.strictEqual(result, 3, 'Month(Date()) = 3 (March)');
    console.log('[PASS] Month(Date()) =', result);
}

// 3. Day() from mocked Now
{
    const ev = makeEv(() => FIXED);
    const result = run(ev, `
Function Test3() As Long
    Test3 = Day(Now())
End Function
`, 'Test3');
    assert.strictEqual(result, 15, 'Day(Now()) = 15');
    console.log('[PASS] Day(Now()) =', result);
}

// 4. Hour/Minute/Second from mocked Now
{
    const ev = makeEv(() => FIXED);
    const result = run(ev, `
Function Test4() As String
    Test4 = Hour(Now()) & ":" & Minute(Now()) & ":" & Second(Now())
End Function
`, 'Test4');
    assert.strictEqual(result, '10:30:45', 'Hour:Minute:Second');
    console.log('[PASS] Hour:Minute:Second =', result);
}

// 5. Time returns fractional day (time-only part)
{
    const ev = makeEv(() => FIXED);
    const result = run(ev, `
Function Test5() As Long
    Test5 = Hour(Time())
End Function
`, 'Test5');
    assert.strictEqual(result, 10, 'Hour(Time()) = 10');
    console.log('[PASS] Hour(Time()) =', result);
}

// 6. VBA branch based on mocked year
{
    const ev = makeEv(() => new Date(2024, 0, 1)); // 2024-01-01 local
    const result = run(ev, `
Function Test6() As String
    If Year(Now()) = 2024 Then
        Test6 = "leap year"
    Else
        Test6 = "other"
    End If
End Function
`, 'Test6');
    assert.strictEqual(result, 'leap year', 'Branch on mocked year');
    console.log('[PASS] Branch on mocked year:', result);
}

// 7. mockDate changes result mid-test via setNowFn
{
    const ev = new Evaluator(() => {});
    const code = `
Function GetYear() As Long
    GetYear = Year(Now())
End Function
`;
    const ast = new Parser(new Lexer(code).tokenize()).parse();
    ev.evaluate(ast);

    ev.setNowFn(() => new Date(2020, 5, 1)); // 2020-06-01 local
    const r1 = ev.callProcedure('GetYear', []);
    assert.strictEqual(r1, 2020, 'First mock: 2020');

    ev.setNowFn(() => new Date(2030, 10, 11)); // 2030-11-11 local
    const r2 = ev.callProcedure('GetYear', []);
    assert.strictEqual(r2, 2030, 'Second mock: 2030');

    ev.setNowFn(null);
    const r3 = ev.callProcedure('GetYear', []);
    assert.ok(r3 >= 2024, 'After clearMock: real year >= 2024');
    console.log('[PASS] setNowFn changes dynamically');
}

// 8. VBARunner.mockDate integration
{
    const vt = new VBARunner('/dev/null');
    // Load code directly
    const ast = new Parser(new Lexer(`
Function GetMonth() As Long
    GetMonth = Month(Now())
End Function
`).tokenize()).parse();
    vt.evaluator.evaluate(ast);

    vt.mockDate('2024-07-04T12:00:00Z');
    const month = vt.evaluator.callProcedure('GetMonth', []);
    assert.strictEqual(month, 7, 'VBARunner.mockDate: July');

    vt.mockDate(null);
    const realMonth = vt.evaluator.callProcedure('GetMonth', []);
    assert.ok(realMonth >= 1 && realMonth <= 12, 'After mockDate(null): valid month');
    console.log('[PASS] VBARunner.mockDate integration');
}

console.log('\n✅ Time Mocking: 全テスト通過');
