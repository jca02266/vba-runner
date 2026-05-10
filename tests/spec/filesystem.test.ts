import { Evaluator } from '../../src/compiler/evaluator';
import { Lexer } from '../../src/compiler/lexer';
import { Parser } from '../../src/compiler/parser';
import * as fs from 'fs';
import * as path from 'path';

function testFileSystem() {
    console.log("Running FileSystem tests...");

    const sandboxRoot = path.join(process.cwd(), "tests/sandbox");
    if (!fs.existsSync(sandboxRoot)) fs.mkdirSync(sandboxRoot, { recursive: true });

    const tempFile = path.join(sandboxRoot, "vba_test_file.txt");
    const copyFile = path.join(sandboxRoot, "vba_test_file_copy.txt");

    // Prepare temp file
    fs.writeFileSync(tempFile, "Hello VBA FileSystem");

    const tests = [
        { code: "Debug.Print CurDir()", expected: "\\" },
        { code: "Debug.Print Dir(\"vba_test_file.txt\")", expected: "vba_test_file.txt" },
        { code: "Debug.Print FileLen(\"vba_test_file.txt\")", expected: String(fs.statSync(tempFile).size) },
        { 
            code: "FileCopy \"vba_test_file.txt\", \"vba_test_file_copy.txt\"\nDebug.Print Dir(\"vba_test_file_copy.txt\")", 
            expected: "vba_test_file_copy.txt" 
        },
        {
            code: "Kill \"vba_test_file.txt\"\nKill \"vba_test_file_copy.txt\"\nDebug.Print Dir(\"vba_test_file.txt\")",
            expected: ""
        }
    ];

    let passed = 0;
    for (const t of tests) {
        let output = "";
        const lexer = new Lexer(t.code);
        const tokens = lexer.tokenize();
        const parser = new Parser(tokens);
        const program = parser.parse();
        const evaluator = new Evaluator((o) => { 
            if (output) output += "\n";
            output += o; 
        }, { sandboxRoot });
        evaluator.evaluate(program);

        const lines = output.trim().split("\n");
        const actual = lines[lines.length - 1]; // Get last output line
        if (actual === t.expected) {
            passed++;
        } else {
            console.log(`[FAIL] ${t.code} - Expected ${t.expected} but got ${actual}`);
        }
    }

    // Cleanup
    if (fs.existsSync(tempFile)) fs.unlinkSync(tempFile);
    if (fs.existsSync(copyFile)) fs.unlinkSync(copyFile);

    console.log(`\nFileSystem Tests: ${passed}/${tests.length} passed`);
    if (passed === tests.length) {
        console.log("✅ FileSystem: 全テスト通過");
    } else {
        process.exit(1);
    }
}

testFileSystem();
