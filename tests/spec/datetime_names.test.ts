import { Evaluator } from '../../src/engine/evaluator';
import { Lexer } from '../../src/engine/lexer';
import { Parser } from '../../src/engine/parser';

function runVba(code: string): any {
    const lexer = new Lexer(code);
    const tokens = lexer.tokenize();
    const parser = new Parser(tokens);
    const program = parser.parse();
    const evaluator = new Evaluator((output) => {
        // console.log(output);
    });
    return evaluator.evaluate(program);
}

function testDateTimeNames() {
    console.log("Running DateTime Name tests...");

    const tests = [
        { code: "Debug.Print MonthName(1)", expected: "January" },
        { code: "Debug.Print MonthName(12, True)", expected: "Dec" },
        { code: "Debug.Print WeekdayName(1)", expected: "Sunday" },
        { code: "Debug.Print WeekdayName(1, True)", expected: "Sun" },
        { code: "Debug.Print WeekdayName(1, False, 2)", expected: "Monday" }, // firstDayOfWeek=2 (Monday), so day 1 is Monday
        { code: "Debug.Print WeekdayName(7, False, 1)", expected: "Saturday" },
    ];

    let passed = 0;
    for (const t of tests) {
        let output = "";
        const lexer = new Lexer(t.code);
        const tokens = lexer.tokenize();
        const parser = new Parser(tokens);
        const program = parser.parse();
        const evaluator = new Evaluator((o) => { output = o; });
        evaluator.evaluate(program);

        const actual = output.trim();
        if (actual === t.expected) {
            passed++;
        } else {
            console.log(`[FAIL] ${t.code} - Expected ${t.expected} but got ${actual}`);
        }
    }

    console.log(`\nDateTime Name Tests: ${passed}/${tests.length} passed`);
    if (passed === tests.length) {
        console.log("✅ DateTime Names: 全テスト通過");
    } else {
        process.exit(1);
    }
}

testDateTimeNames();
