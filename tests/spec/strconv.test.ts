import { Lexer } from '../../src/engine/lexer';
import { Parser } from '../../src/engine/parser';
import { Evaluator } from '../../src/engine/evaluator';
import { assert } from '../../test-libs/test-runner';

// StrConv tests
{
    const code = `
    Sub Test()
        Debug.Print StrConv("hello WORLD", 1) ' vbUpperCase
        Debug.Print StrConv("hello WORLD", 2) ' vbLowerCase
        Debug.Print StrConv("hello world", 3) ' vbProperCase
        Debug.Print StrConv("apple pie. orange juice", 3) ' vbProperCase
        
        ' Japanese tests (optional but good to have)
        ' Debug.Print StrConv("あいうえお", 16) ' vbKatakana
    End Sub
    `;
    let output = "";
    const ev = new Evaluator(s => output += s + "\n");
    const tokens = new Lexer(code).tokenize();
    const ast = new Parser(tokens).parse();
    ev.evaluate(ast);
    ev.callProcedure("Test", []);
    
    const lines = output.trim().split("\n");
    assert.strictEqual(lines[0], "HELLO WORLD", "StrConv(..., 1) -> UpperCase");
    assert.strictEqual(lines[1], "hello world", "StrConv(..., 2) -> LowerCase");
    assert.strictEqual(lines[2], "Hello World", "StrConv(..., 3) -> ProperCase");
    assert.strictEqual(lines[3], "Apple Pie. Orange Juice", "StrConv(..., 3) -> ProperCase (complex)");
    console.log('[PASS] StrConv basic tests');
}

// Japanese StrConv tests
{
    const code = `
    Sub Test()
        ' vbKatakana = 16, vbHiragana = 32
        Debug.Print StrConv("あいうえお", 16)
        Debug.Print StrConv("アイウエオ", 32)
        ' vbWide = 4, vbNarrow = 8
        Debug.Print StrConv("ABC", 4)
        Debug.Print StrConv("ＡＢＣ", 8)
    End Sub
    `;
    let output = "";
    const ev = new Evaluator(s => output += s + "\n");
    const tokens = new Lexer(code).tokenize();
    const ast = new Parser(tokens).parse();
    ev.evaluate(ast);
    ev.callProcedure("Test", []);
    
    const lines = output.trim().split("\n");
    assert.strictEqual(lines[0], "アイウエオ", "StrConv(..., 16) -> Katakana");
    assert.strictEqual(lines[1], "あいうえお", "StrConv(..., 32) -> Hiragana");
    assert.strictEqual(lines[2], "ＡＢＣ", "StrConv(..., 4) -> Wide");
    assert.strictEqual(lines[3], "ABC", "StrConv(..., 8) -> Narrow");
    console.log('[PASS] StrConv Japanese tests');
}

console.log('\n✅ StrConv: 全テスト通過');
