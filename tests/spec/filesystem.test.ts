import { Evaluator } from '../../src/compiler/evaluator';
import { Lexer } from '../../src/compiler/lexer';
import { Parser } from '../../src/compiler/parser';
import * as fs from 'fs';
import * as path from 'path';

function testFileSystem() {
    console.log("Running FileSystem tests...");

    // Setup: clean sandbox/test.txt
    const sandbox = path.resolve(process.cwd(), 'sandbox');
    if (!fs.existsSync(sandbox)) fs.mkdirSync(sandbox);
    const testFile = path.join(sandbox, 'test.txt');
    if (fs.existsSync(testFile)) fs.unlinkSync(testFile);

    const vbaCode = `
        Sub TestFile()
            Dim fn As Integer
            fn = FreeFile()
            Open "test.txt" For Output As #fn
            Print #fn, "Hello VBA"
            Print #fn, "Line 2"
            Close #fn
            
            Dim lineStr As String
            Open "test.txt" For Input As #1
            Line Input #1, lineStr
            Debug.Print lineStr
            Line Input #1, lineStr
            Debug.Print lineStr
            Close #1
            
            Kill "test.txt"
            
            ' Environ test (mocked)
            Debug.Print Environ("TEMP")
        End Sub
        TestFile
    `;

    let output = "";
    const lexer = new Lexer(vbaCode);
    const tokens = lexer.tokenize();
    const parser = new Parser(tokens);
    const program = parser.parse();
    // Setup initial env for Sandbox
    const evaluator = new Evaluator((o) => { output += o + "\n"; }, { env: { "temp": "/tmp/vba" } });
    evaluator.evaluate(program);

    const lines = output.trim().split("\n").map(l => l.trim());
    const expected = [
        "Hello VBA",
        "Line 2",
        "/tmp/vba"
    ];

    let passed = 0;
    for (let i = 0; i < expected.length; i++) {
        if (lines[i] === expected[i]) {
            passed++;
        } else {
            console.log(`[FAIL] FileSystem line ${i+1} - Expected "${expected[i]}" but got "${lines[i]}"`);
        }
    }

    if (passed === expected.length) {
        console.log("✅ FileSystem: 全テスト通過");
    } else {
        console.log("Actual output lines:", lines);
        process.exit(1);
    }
}

testFileSystem();
