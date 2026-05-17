import { Lexer } from '../../src/engine/lexer';
import { Parser } from '../../src/engine/parser';
import { Evaluator } from '../../src/engine/evaluator';
import { assert } from '../../test-libs/test-runner';

function runFunc(code: string, name: string, args: any[] = []): any {
    const tokens = new Lexer(code).tokenize();
    const ast = new Parser(tokens).parse();
    const ev = new Evaluator(() => {});
    ev.evaluate(ast);
    return ev.callProcedure(name, args);
}

// 1. Mid Statement (Assignment)
{
    const code = `
        Function TestMid()
            Dim s As String
            s = "The dog jumps"
            Mid(s, 5, 3) = "fox"
            TestMid = s
        End Function
    `;
    assert.strictEqual(runFunc(code, 'TestMid'), "The fox jumps", 'Mid Statement with length');
}

{
    const code = `
        Function TestMidNoLen()
            Dim s As String
            s = "The dog jumps"
            Mid(s, 5) = "cat"
            TestMidNoLen = s
        End Function
    `;
    assert.strictEqual(runFunc(code, 'TestMidNoLen'), "The cat jumps", 'Mid Statement without length');
}

{
    const code = `
        Function TestMidLonger()
            Dim s As String
            s = "The dog"
            Mid(s, 5) = "quick brown fox"
            TestMidLonger = s
        End Function
    `;
    // VBA Mid statement: "the number of characters replaced is always less than or equal to the length of stringvar"
    // So "The dog" (7 chars). Mid(s, 5) = "quick..." -> replaces from index 5 (3 chars left).
    // Result: "The qui"
    assert.strictEqual(runFunc(code, 'TestMidLonger'), "The qui", 'Mid Statement with longer source string');
}

console.log('\n✅ Mid Statement: 全テスト通過');
