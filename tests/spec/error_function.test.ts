import { Lexer } from '../../src/compiler/lexer';
import { Parser } from '../../src/compiler/parser';
import { Evaluator } from '../../src/compiler/evaluator';
import { assert } from '../ts/test-runner';

// Error function tests
{
    const code = `
    Sub Test()
        Debug.Print Error(6)
        Debug.Print Error(11)
        Debug.Print Error(13)
        Debug.Print Error(999) ' Unknown
        
        ' Without arguments (uses most recent error)
        On Error Resume Next
        Err.Raise 13
        Debug.Print Error
    End Sub
    `;
    let output = "";
    const ev = new Evaluator(s => output += s + "\n");
    const tokens = new Lexer(code).tokenize();
    const ast = new Parser(tokens).parse();
    ev.evaluate(ast);
    ev.callProcedure("Test", []);
    
    const lines = output.trim().split("\n");
    assert.strictEqual(lines[0], "Overflow", "Error(6)");
    assert.strictEqual(lines[1], "Division by zero", "Error(11)");
    assert.strictEqual(lines[2], "Type mismatch", "Error(13)");
    assert.strictEqual(lines[3], "Application-defined or object-defined error", "Error(999)");
    assert.strictEqual(lines[4], "Type mismatch", "Error (without args)");
    console.log('[PASS] Error function tests');
}

console.log('\n✅ Error Function: 全テスト通過');
