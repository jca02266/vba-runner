import { Evaluator } from '../../src/compiler/evaluator';
import { Lexer } from '../../src/compiler/lexer';
import { Parser } from '../../src/compiler/parser';

function testShell() {
    console.log("Running Shell tests...");

    // Since Shell is async and might depend on environment, we just check if it returns a pid-like number
    const code = "Debug.Print Shell(\"echo hello\") > 0";
    
    let output = "";
    const lexer = new Lexer(code);
    const tokens = lexer.tokenize();
    const parser = new Parser(tokens);
    const program = parser.parse();
    const evaluator = new Evaluator((o) => { output = o; });
    evaluator.evaluate(program);

    const actual = output.trim();
    if (actual === "True") {
        console.log("✅ Shell: テスト通過");
    } else {
        console.log(`[FAIL] Shell - Expected True but got ${actual}`);
        process.exit(1);
    }
}

testShell();
