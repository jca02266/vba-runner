import { Evaluator } from '../../src/compiler/evaluator';
import { Lexer } from '../../src/compiler/lexer';
import { Parser } from '../../src/compiler/parser';

function testCDec() {
    console.log("Running CDec tests...");

    const tests = [
        { code: "Debug.Print CDec(123.45)", expected: "123.45" },
        { code: "Debug.Print TypeName(CDec(1))", expected: "Decimal" },
        { code: "Debug.Print VarType(CDec(1))", expected: "14" },
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

    console.log(`\nCDec Tests: ${passed}/${tests.length} passed`);
    if (passed === tests.length) {
        console.log("✅ CDec: 全テスト通過");
    } else {
        process.exit(1);
    }
}

testCDec();
