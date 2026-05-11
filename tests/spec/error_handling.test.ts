import { Evaluator } from '../../src/compiler/evaluator';
import { Lexer } from '../../src/compiler/lexer';
import { Parser } from '../../src/compiler/parser';

function testErrorHandling() {
    console.log("Running Error Handling tests...");

    const vbaCode = `
        Sub TestError()
            On Error GoTo ErrorHandler
            Debug.Print "Start"
            Err.Raise 513, "Test", "Custom Error"
            Debug.Print "Back from Error"
            GoTo Done
        ErrorHandler:
            Debug.Print "Caught: " & Err.Number
            Resume Next
        Done:
            Debug.Print "Finished"
        End Sub
        TestError
    `;

    let output = "";
    const lexer = new Lexer(vbaCode);
    const tokens = lexer.tokenize();
    const parser = new Parser(tokens);
    const program = parser.parse();
    const evaluator = new Evaluator((o) => { output += o + "\n"; });
    evaluator.evaluate(program);

    const lines = output.trim().split("\n").map(l => l.trim());
    const expected = [
        "Start",
        "Caught: 513",
        "Back from Error",
        "Finished"
    ];

    let passed = 0;
    for (let i = 0; i < expected.length; i++) {
        if (lines[i] === expected[i]) {
            passed++;
        } else {
            console.log(`[FAIL] Error Handling line ${i+1} - Expected "${expected[i]}" but got "${lines[i]}"`);
        }
    }

    if (passed === expected.length) {
        console.log("✅ Error Handling: 全テスト通過");
    } else {
        console.log("Actual output lines:", lines);
        process.exit(1);
    }
}

testErrorHandling();
