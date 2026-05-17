import { Evaluator } from '../../src/engine/evaluator';
import { Lexer } from '../../src/engine/lexer';
import { Parser } from '../../src/engine/parser';

function testStrFix() {
    console.log("Running Str fixes tests...");

    const tests = [
        { code: "Debug.Print Str(123)", expected: " 123" },
        { code: "Debug.Print Str(-123)", expected: "-123" },
        { code: "Debug.Print Str(0)", expected: " 0" },
        { code: "Debug.Print Str(Null)", expected: "Null" },
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

        const actual = output;
        if (actual === t.expected) {
            passed++;
        } else {
            console.log(`[FAIL] ${t.code} - Expected '${t.expected}' but got '${actual}'`);
        }
    }

    console.log(`\nStr Fixes Tests: ${passed}/${tests.length} passed`);
    if (passed === tests.length) {
        console.log("✅ Str Fixes: 全テスト通過");
    } else {
        process.exit(1);
    }
}

testStrFix();
