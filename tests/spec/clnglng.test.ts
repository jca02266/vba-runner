import { Evaluator } from '../../src/engine/evaluator';
import { Lexer } from '../../src/engine/lexer';
import { Parser } from '../../src/engine/parser';

function testCLngLng() {
    console.log("Running CLngLng tests...");

    const tests = [
        { code: "Debug.Print CLngLng(123.45)", expected: "123" },
        { code: "Debug.Print CLngLng(-123.45)", expected: "-123" },
        { code: "Debug.Print CLngLng(True)", expected: "-1" },
        { code: "Debug.Print CLngLng(\"9223372036854775807\")", expected: "9223372036854775807" },
        { code: "Debug.Print CLngLng(\"-9223372036854775808\")", expected: "-9223372036854775808" },
        { code: "Debug.Print TypeName(CLngLng(1))", expected: "LongLong" },
        { code: "Debug.Print VarType(CLngLng(1))", expected: "20" },
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

    console.log(`\nCLngLng Tests: ${passed}/${tests.length} passed`);
    if (passed === tests.length) {
        console.log("✅ CLngLng: 全テスト通過");
    } else {
        process.exit(1);
    }
}

testCLngLng();
