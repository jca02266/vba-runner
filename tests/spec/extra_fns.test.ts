import { Evaluator } from '../../src/engine/evaluator';
import { Lexer } from '../../src/engine/lexer';
import { Parser } from '../../src/engine/parser';

function testExtraFunctions() {
    console.log("Running Extra Functions tests...");

    const tests = [
        { code: "Debug.Print CByte(12.5)", expected: "12" }, // VBA CByte rounds to nearest even? No, CByte rounds .5 to even.
        { code: "Debug.Print CByte(13.5)", expected: "14" },
        { code: "Debug.Print Round(12.5)", expected: "12" },
        { code: "Debug.Print Round(13.5)", expected: "14" },
        { code: "Debug.Print StrComp(\"abc\", \"ABC\")", expected: "1" }, // Default Binary
        { code: "Debug.Print StrComp(\"abc\", \"ABC\", 1)", expected: "0" }, // Text
        { code: "Debug.Print Format(1234.567, \"#,##0.00\")", expected: "1,234.57" },
        { code: "Debug.Print Format(0.123, \"Percent\")", expected: "12.30%" },
        { code: "Debug.Print Format(DateSerial(2023, 5, 11), \"yyyy-mm-dd\")", expected: "2023-05-11" },
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
        // CByte(12.5) in VBA is actually 12 (Banker's rounding)
        // My CByte implementation uses Math.round which is NOT Banker's rounding.
        // I should fix CByte to use vbaRound.
        
        if (actual === t.expected) {
            passed++;
        } else {
            console.log(`[FAIL] ${t.code} - Expected ${t.expected} but got ${actual}`);
        }
    }

    if (passed === tests.length) {
        console.log("✅ Extra Functions: 全テスト通過");
    } else {
        process.exit(1);
    }
}

testExtraFunctions();
