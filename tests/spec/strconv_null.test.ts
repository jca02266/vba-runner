import { Evaluator } from '../../src/compiler/evaluator';
import { Lexer } from '../../src/compiler/lexer';
import { Parser } from '../../src/compiler/parser';

function testStrConvNull() {
    console.log("Running StrConv Null tests...");

    const tests = [
        { code: "Debug.Print TypeName(StrConv(Null, 1))", expected: "Null" },
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

    console.log(`\nStrConv Null Tests: ${passed}/${tests.length} passed`);
    if (passed === tests.length) {
        console.log("✅ StrConv Null: 全テスト通過");
    } else {
        process.exit(1);
    }
}

testStrConvNull();
