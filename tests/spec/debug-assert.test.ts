import { Lexer } from '../../src/compiler/lexer';
import { Parser } from '../../src/compiler/parser';
import { Evaluator } from '../../src/compiler/evaluator';
import { assert as specAssert } from '../ts/test-runner';

// Debug.Assert tests
{
    const code = `
    Sub Test()
        Debug.Assert True
        Debug.Print "Pass1"
        ' Debug.Assert False ' This should throw or stop
    End Sub
    `;
    let output = "";
    const ev = new Evaluator(s => output += s + "\n");
    const tokens = new Lexer(code).tokenize();
    const ast = new Parser(tokens).parse();
    ev.evaluate(ast);
    ev.callProcedure("Test", []);
    
    const lines = output.trim().split("\n");
    specAssert.strictEqual(lines[0], "Pass1", "Debug.Assert True should pass");
    console.log('[PASS] Debug.Assert True');
}

{
    const code = `
    Sub Test()
        On Error Resume Next
        Debug.Assert False
        If Err.Number <> 0 Then Debug.Print "Error:" & Err.Description
    End Sub
    `;
    let output = "";
    const ev = new Evaluator(s => output += s + "\n");
    const tokens = new Lexer(code).tokenize();
    const ast = new Parser(tokens).parse();
    ev.evaluate(ast);
    ev.callProcedure("Test", []);
    
    const lines = output.trim().split("\n");
    specAssert.strictEqual(lines[0], "Error:Assertion failed", "Debug.Assert False should raise an error");
    console.log('[PASS] Debug.Assert False');
}

console.log('\n✅ Debug.Assert: 全テスト通過');
