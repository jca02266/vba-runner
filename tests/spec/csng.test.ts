import { Lexer } from '../../src/compiler/lexer';
import { Parser } from '../../src/compiler/parser';
import { Evaluator } from '../../src/compiler/evaluator';
import { assert } from '../ts/test-runner';

// CSng tests
{
    const code = `
    Sub Test()
        Debug.Print CSng(1.23456789)
        Debug.Print CSng(123456.7)
        Debug.Print CSng(True)
        Debug.Print CSng("1.5")
    End Sub
    `;
    let output = "";
    const ev = new Evaluator(s => output += s + "\n");
    const tokens = new Lexer(code).tokenize();
    const ast = new Parser(tokens).parse();
    ev.evaluate(ast);
    ev.callProcedure("Test", []);
    
    const lines = output.trim().split("\n");
    // CSng(1.23456789) -> 1.2345678806304932 (float32 representation)
    // When printed, it might be truncated or showing more digits.
    // In VBA, it usually shows ~7 digits.
    assert.strictEqual(Math.abs(parseFloat(lines[0]) - 1.2345679) < 0.000001, true, "CSng(1.23456789)");
    assert.strictEqual(Math.abs(parseFloat(lines[1]) - 123456.7) < 0.1, true, "CSng(123456.7)");
    assert.strictEqual(lines[2], "-1", "CSng(True)");
    assert.strictEqual(lines[3], "1.5", "CSng('1.5')");
    console.log('[PASS] CSng basic tests');
}

// Overflow tests
{
    const code = `
    Sub Test()
        On Error Resume Next
        ' Max Single is approx 3.4E38
        Debug.Print CSng(4E+38)
        If Err.Number <> 0 Then Debug.Print "Error:" & Err.Number
    End Sub
    `;
    let output = "";
    const ev = new Evaluator(s => output += s + "\n");
    const tokens = new Lexer(code).tokenize();
    const ast = new Parser(tokens).parse();
    ev.evaluate(ast);
    ev.callProcedure("Test", []);
    
    const lines = output.trim().split("\n");
    assert.strictEqual(lines[0], "Error:6", "CSng(4E+38) should overflow");
    console.log('[PASS] CSng Overflow tests');
}

console.log('\n✅ CSng: 全テスト通過');
