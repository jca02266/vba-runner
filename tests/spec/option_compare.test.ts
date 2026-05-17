import { assert } from '../../test-libs/test-runner';
import { Evaluator } from '../../src/engine/evaluator';
import { Parser } from '../../src/engine/parser';
import { Lexer } from '../../src/engine/lexer';

function evalVBA(code: string) {
    const tokens = new Lexer(code).tokenize();
    const parser = new Parser(tokens);
    const program = parser.parse();
    const evaluator = new Evaluator((s) => console.log(s));
    evaluator.evaluate(program);
    return evaluator;
}

console.log('[Test Suite] Option Compare の検証');

// Case 1: Default (Binary)
const codeBinary = `
    Function TestBinary()
        TestBinary = ("abc" = "ABC")
    End Function
    Function TestLikeBinary()
        TestLikeBinary = ("abc" Like "A*")
    End Function
`;
const evBinary = evalVBA(codeBinary);
assert.strictEqual(evBinary.callProcedure('TestBinary', []).value, 0, 'Default Binary: "abc" = "ABC" should be False');
assert.strictEqual(evBinary.callProcedure('TestLikeBinary', []).value, 0, 'Default Binary: "abc" Like "A*" should be False');

// Case 2: Option Compare Text
const codeText = `
    Option Compare Text
    Function TestText()
        TestText = ("abc" = "ABC")
    End Function
    Function TestLikeText()
        TestLikeText = ("abc" Like "A*")
    End Function
    Function TestInStrText()
        TestInStrText = InStr("abc", "B")
    End Function
`;
const evText = evalVBA(codeText);
assert.strictEqual(evText.callProcedure('TestText', []).value, -1, 'Option Compare Text: "abc" = "ABC" should be True');
assert.strictEqual(evText.callProcedure('TestLikeText', []).value, -1, 'Option Compare Text: "abc" Like "A*" should be True');
assert.strictEqual(evText.callProcedure('TestInStrText', []), 2, 'Option Compare Text: InStr("abc", "B") should be 2');

console.log('✅ Option Compare: 全テスト通過');
