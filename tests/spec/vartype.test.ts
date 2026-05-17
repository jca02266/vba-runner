import { Evaluator } from '../../src/engine/evaluator';
import { Lexer } from '../../src/engine/lexer';
import { Parser } from '../../src/engine/parser';
import { VbaErrorValue } from '../../src/engine/evaluator';

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

function testVarType() {
    console.log("Running VarType tests...");

    const tests = [
        { code: "Debug.Print VarType(Empty)", expected: 0 },
        { code: "Debug.Print VarType(Null)", expected: 1 },
        { code: "Debug.Print VarType(True)", expected: 11 },
        { code: "Debug.Print VarType(#2023-01-01#)", expected: 7 },
        { code: "Debug.Print VarType(CVErr(5))", expected: 10 },
        { code: "Debug.Print VarType(1.23)", expected: 5 },
        { code: "Debug.Print VarType(\"hello\")", expected: 8 },
        { code: "Dim a(1): Debug.Print VarType(a)", expected: 8192 + 12 },
        { code: "Debug.Print vbInteger", expected: 2 },
        { code: "Debug.Print vbLong", expected: 3 },
        { code: "Debug.Print vbSingle", expected: 4 },
        { code: "Debug.Print vbDouble", expected: 5 },
        { code: "Debug.Print vbBoolean", expected: 11 },
        { code: "Debug.Print vbArray", expected: 8192 },
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

        const actual = Number(output.trim());
        if (actual === t.expected) {
            passed++;
        } else {
            console.log(`[FAIL] ${t.code} - Expected ${t.expected} but got ${actual}`);
        }
    }

    console.log(`\nVarType Tests: ${passed}/${tests.length} passed`);
    if (passed === tests.length) {
        console.log("✅ VarType: 全テスト通過");
    } else {
        process.exit(1);
    }
}

testVarType();
