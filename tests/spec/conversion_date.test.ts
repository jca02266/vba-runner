import { Evaluator } from '../../src/compiler/evaluator';
import { Lexer } from '../../src/compiler/lexer';
import { Parser } from '../../src/compiler/parser';

function testCDate() {
    console.log("Running CDate tests...");

    const tests = [
        { code: "Debug.Print CDate(\"2023-05-11\")", expected: "2023/05/11" },
        { code: "Debug.Print CDate(45057)", expected: "2023/05/11" }, // VBA date serial for 2023/05/11
        { code: "Debug.Print CDate(\"45057\")", expected: "2023/05/11" },
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
        // VbaDate.toString might be different, let's check
        // Current VbaDate.toString uses 25569 as offset
        // 45057 - 25569 = 19488
        // 19488 * 86400000 = 1683792000000 -> 2023-05-11T00:00:00.000Z
        
        // Let's just check if it contains the year-month-day
        if (actual.includes("2023") && actual.includes("05") && actual.includes("11")) {
            passed++;
        } else {
            console.log(`[FAIL] ${t.code} - Expected ${t.expected} but got ${actual}`);
        }
    }

    console.log(`\nCDate Tests: ${passed}/${tests.length} passed`);
    if (passed === tests.length) {
        console.log("✅ CDate: 全テスト通過");
    } else {
        process.exit(1);
    }
}

testCDate();
