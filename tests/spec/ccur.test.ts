import { Lexer } from '../../src/engine/lexer';
import { Parser } from '../../src/engine/parser';
import { Evaluator } from '../../src/engine/evaluator';
import { assert } from '../../test-libs/test-runner';

// CCur tests
{
    const code = `
    Sub Test()
        Debug.Print CCur(10.12344)
        Debug.Print CCur(10.12345)
        Debug.Print CCur(10.12346)
        Debug.Print CCur(10.12355)
        Debug.Print CCur(-1.23455)
        Debug.Print CCur(True)
    End Sub
    `;
    let output = "";
    const ev = new Evaluator(s => output += s + "\n");
    const tokens = new Lexer(code).tokenize();
    const ast = new Parser(tokens).parse();
    ev.evaluate(ast);
    ev.callProcedure("Test", []);
    
    const lines = output.trim().split("\n");
    assert.strictEqual(lines[0], "10.1234", "CCur(10.12344) -> 10.1234");
    assert.strictEqual(lines[1], "10.1234", "CCur(10.12345) -> 10.1234 (Banker's round to even)");
    assert.strictEqual(lines[2], "10.1235", "CCur(10.12346) -> 10.1235");
    assert.strictEqual(lines[3], "10.1236", "CCur(10.12355) -> 10.1236 (Banker's round to even)");
    assert.strictEqual(lines[4], "-1.2346", "CCur(-1.23455) -> -1.2346");
    assert.strictEqual(lines[5], "-1", "CCur(True) -> -1");
    console.log('[PASS] CCur tests');
}

// Overflow tests
{
    const code = `
    Sub Test()
        On Error Resume Next
        ' Max Currency is approx 922 trillion
        Debug.Print CCur(1E+15)
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
    assert.strictEqual(lines[0], "Error:6", "CCur(1E+15) should overflow");
    console.log('[PASS] CCur Overflow tests');
}

console.log('\n✅ CCur: 全テスト通過');
