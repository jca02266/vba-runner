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

function testEnviron() {
    console.log("Running Environ tests...");

    const tests = [
        { code: "Debug.Print Environ(\"VBA_TEST_VAR\")", expected: "HelloWorld" },
        { code: "Debug.Print Environ(\"NON_EXISTENT_VAR\")", expected: "" },
        { code: "Debug.Print Environ$(1)", expected: "VBA_TEST_VAR=HelloWorld" },
        { code: "Debug.Print Environ(9999)", expected: "" },
    ];

    let passed = 0;
    for (const t of tests) {
        let output = "";
        const lexer = new Lexer(t.code);
        const tokens = lexer.tokenize();
        const parser = new Parser(tokens);
        const program = parser.parse();
        const evaluator = new Evaluator((o) => { output = o; }, {
            env: { VBA_TEST_VAR: "HelloWorld" }
        });
        evaluator.evaluate(program);

        const actual = output.trim();
        if (actual === t.expected || (t.expected === "True" && actual === "True")) {
            passed++;
        } else {
            console.log(`[FAIL] ${t.code} - Expected ${t.expected} but got ${actual}`);
        }
    }

    console.log(`\nEnviron Tests: ${passed}/${tests.length} passed`);
    if (passed === tests.length) {
        console.log("✅ Environ: 全テスト通過");
    } else {
        process.exit(1);
    }
}

testEnviron();
