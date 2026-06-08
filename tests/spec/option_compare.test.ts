import { evalVBASingle, assert } from '../../test-libs/test-runner';

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
const evBinary = evalVBASingle(codeBinary);
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
const evText = evalVBASingle(codeText);
assert.strictEqual(evText.callProcedure('TestText', []).value, -1, 'Option Compare Text: "abc" = "ABC" should be True');
assert.strictEqual(evText.callProcedure('TestLikeText', []).value, -1, 'Option Compare Text: "abc" Like "A*" should be True');
assert.strictEqual(evText.callProcedure('TestInStrText', []), 2, 'Option Compare Text: InStr("abc", "B") should be 2');

console.log('✅ Option Compare: 全テスト通過');
