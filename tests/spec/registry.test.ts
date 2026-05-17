import { Evaluator } from '../../src/engine/evaluator';
import { Lexer } from '../../src/engine/lexer';
import { Parser } from '../../src/engine/parser';

function testRegistry() {
    console.log("Running Registry tests...");

    const vbaCode = `
        SaveSetting "MyApp", "Section1", "Key1", "Value1"
        Debug.Print GetSetting("MyApp", "Section1", "Key1")
        SaveSetting "MyApp", "Section1", "Key1", "NewValue"
        Debug.Print GetSetting("MyApp", "Section1", "Key1")
        Debug.Print GetSetting("MyApp", "Section1", "Key2", "Default")
        DeleteSetting "MyApp", "Section1", "Key1"
        Debug.Print GetSetting("MyApp", "Section1", "Key1", "Deleted")
    `;

    let output = "";
    const lexer = new Lexer(vbaCode);
    const tokens = lexer.tokenize();
    const parser = new Parser(tokens);
    const program = parser.parse();
    const evaluator = new Evaluator((o) => { output += o + "\n"; });
    evaluator.evaluate(program);

    const lines = output.trim().split("\n");
    const expected = [
        "Value1",
        "NewValue",
        "Default",
        "Deleted"
    ];

    let passed = 0;
    for (let i = 0; i < expected.length; i++) {
        if (lines[i].trim() === expected[i]) {
            passed++;
        } else {
            console.log(`[FAIL] Registry line ${i+1} - Expected ${expected[i]} but got ${lines[i]}`);
        }
    }

    if (passed === expected.length) {
        console.log("✅ Registry (Settings): 全テスト通過");
    } else {
        process.exit(1);
    }
}

testRegistry();
