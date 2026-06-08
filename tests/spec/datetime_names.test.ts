import { evalVBASingle, assert } from '../../test-libs/test-runner';

console.log("Running DateTime Name tests...");

const tests = [
    { code: "Debug.Print MonthName(1)", expected: "January" },
    { code: "Debug.Print MonthName(12, True)", expected: "Dec" },
    { code: "Debug.Print WeekdayName(1)", expected: "Sunday" },
    { code: "Debug.Print WeekdayName(1, True)", expected: "Sun" },
    { code: "Debug.Print WeekdayName(1, False, 2)", expected: "Monday" },
    { code: "Debug.Print WeekdayName(7, False, 1)", expected: "Saturday" },
];

for (const t of tests) {
    let output = '';
    evalVBASingle(t.code, { onPrint: (o) => { output = o; } });
    assert.strictEqual(output.trim(), t.expected, t.code);
    console.log(`[PASS] ${t.code}`);
}

console.log('✅ DateTime Names: 全テスト通過');
