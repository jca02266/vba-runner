import { Evaluator } from '../../src/engine/evaluator';
import { Lexer } from '../../src/engine/lexer';
import { Parser } from '../../src/engine/parser';

function testFreeFile() {
    console.log("Running FreeFile tests...");

    const tests = [
        { code: "Debug.Print FreeFile()", expected: "1" },
        { code: "Debug.Print FreeFile(1)", expected: "256" },
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

    console.log(`\nFreeFile Tests: ${passed}/${tests.length} passed`);
    if (passed === tests.length) {
        console.log("✅ FreeFile: 全テスト通過");
    } else {
        process.exit(1);
    }
}

testFreeFile();
