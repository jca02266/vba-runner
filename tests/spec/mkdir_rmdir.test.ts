import { Evaluator } from '../../src/engine/evaluator';
import { Lexer } from '../../src/engine/lexer';
import { Parser } from '../../src/engine/parser';
import * as fs from 'fs';

function testMkDirRmDir() {
    console.log("Running MkDir/RmDir tests...");

    const testDir = "vba_test_dir";
    if (fs.existsSync(testDir)) fs.rmdirSync(testDir);

    const code = `
        MkDir "vba_test_dir"
        Debug.Print Dir("vba_test_dir")
        RmDir "vba_test_dir"
    `;

    let output = "";
    const lexer = new Lexer(code);
    const tokens = lexer.tokenize();
    const parser = new Parser(tokens);
    const program = parser.parse();
    const evaluator = new Evaluator((o) => { output = o; });
    evaluator.evaluate(program);

    if (output.trim() === "vba_test_dir" && !fs.existsSync(testDir)) {
        console.log("✅ MkDir/RmDir: 全テスト通過");
    } else {
        console.log(`[FAIL] Expected vba_test_dir and directory removed, but got ${output.trim()}`);
        process.exit(1);
    }
}

testMkDirRmDir();
