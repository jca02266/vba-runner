import { Lexer } from '../../src/engine/lexer';
import { Parser } from '../../src/engine/parser';
import { Evaluator } from '../../src/engine/evaluator';
import { assert, vbaTrue, vbaFalse } from '../../test-libs/test-runner';

function evalVBA(code: string): any {
    const tokens = new Lexer(code).tokenize();
    const ast = new Parser(tokens).parse();
    const ev = new Evaluator(console.log);
    ev.evaluate(ast);
    return ev;
}

function runFunc(code: string, name: string, args: any[] = []): any {
    return evalVBA(code).callProcedure(name, args);
}

// Test 1: Basic character class - single characters
{
    const code = `
    Function Test1()
        Test1 = "a" Like "[abc]"
    End Function
    `;
    const result = runFunc(code, 'Test1');
    assert.strictEqual(result, vbaTrue, '[abc] should match a');
    console.log('[PASS] Test 1: Basic character class single character');
}

// Test 2: Character class - character not in list
{
    const code = `
    Function Test2()
        Test2 = "d" Like "[abc]"
    End Function
    `;
    const result = runFunc(code, 'Test2');
    assert.strictEqual(result, vbaFalse, '[abc] should not match d');
    console.log('[PASS] Test 2: Character not in list');
}

// Test 3: Character range - lowercase letters
{
    const code = `
    Function Test3()
        Test3 = "m" Like "[a-z]"
    End Function
    `;
    const result = runFunc(code, 'Test3');
    assert.strictEqual(result, vbaTrue, '[a-z] should match m');
    console.log('[PASS] Test 3: Range [a-z]');
}

// Test 4: Character range - digits
{
    const code = `
    Function Test4()
        Test4 = "5" Like "[0-9]"
    End Function
    `;
    const result = runFunc(code, 'Test4');
    assert.strictEqual(result, vbaTrue, '[0-9] should match 5');
    console.log('[PASS] Test 4: Range [0-9]');
}

// Test 5: Character range - uppercase letters
{
    const code = `
    Function Test5()
        Test5 = "M" Like "[A-Z]"
    End Function
    `;
    const result = runFunc(code, 'Test5');
    assert.strictEqual(result, vbaTrue, '[A-Z] should match M');
    console.log('[PASS] Test 5: Range [A-Z]');
}

// Test 6: Negated character class
{
    const code = `
    Function Test6()
        Test6 = "d" Like "[!abc]"
    End Function
    `;
    const result = runFunc(code, 'Test6');
    assert.strictEqual(result, vbaTrue, '[!abc] should match d');
    console.log('[PASS] Test 6: Negated class [!abc]');
}

// Test 7: Negated class - character in list
{
    const code = `
    Function Test7()
        Test7 = "a" Like "[!abc]"
    End Function
    `;
    const result = runFunc(code, 'Test7');
    assert.strictEqual(result, vbaFalse, '[!abc] should not match a');
    console.log('[PASS] Test 7: Negated class with char in list');
}

// Test 8: Hyphen at beginning of class
{
    const code = `
    Function Test8()
        Test8 = "-" Like "[-abc]"
    End Function
    `;
    const result = runFunc(code, 'Test8');
    assert.strictEqual(result, vbaTrue, '[-abc] should match hyphen');
    console.log('[PASS] Test 8: Hyphen at start');
}

// Test 9: Hyphen at end of class
{
    const code = `
    Function Test9()
        Test9 = "-" Like "[abc-]"
    End Function
    `;
    const result = runFunc(code, 'Test9');
    assert.strictEqual(result, vbaTrue, '[abc-] should match hyphen at end');
    console.log('[PASS] Test 9: Hyphen at end');
}

// Test 10: Multiple ranges in one class
{
    const code = `
    Function Test10()
        Dim result As Boolean
        result = ("5" Like "[0-9a-z]") And ("m" Like "[0-9a-z]")
        Test10 = result
    End Function
    `;
    const result = runFunc(code, 'Test10');
    assert.strictEqual(result, vbaTrue, '[0-9a-z] should match both 5 and m');
    console.log('[PASS] Test 10: Multiple ranges');
}

// Test 11: Mixed individual chars and ranges
{
    const code = `
    Function Test11()
        Dim result As Boolean
        result = ("a" Like "[a-zA-Z_]") And ("_" Like "[a-zA-Z_]") And ("Z" Like "[a-zA-Z_]")
        Test11 = result
    End Function
    `;
    const result = runFunc(code, 'Test11');
    assert.strictEqual(result, vbaTrue, '[a-zA-Z_] should match a, _, Z');
    console.log('[PASS] Test 11: Mixed ranges and chars');
}

// Test 12: Character class in pattern with other elements
{
    const code = `
    Function Test12()
        Test12 = "test1" Like "test[0-9]"
    End Function
    `;
    const result = runFunc(code, 'Test12');
    assert.strictEqual(result, vbaTrue, 'test[0-9] should match test1');
    console.log('[PASS] Test 12: Class in pattern');
}

// Test 13: Character class with wildcard
{
    const code = `
    Function Test13()
        Test13 = "abc123" Like "[a-z]*[0-9]"
    End Function
    `;
    const result = runFunc(code, 'Test13');
    assert.strictEqual(result, vbaTrue, '[a-z]*[0-9] should match abc123');
    console.log('[PASS] Test 13: Class with wildcard');
}

// Test 14: Empty character class (should be at least empty in [])
{
    const code = `
    Function Test14()
        Dim result As Boolean
        result = "a" Like "[]"
        Test14 = result
    End Function
    `;
    const result = runFunc(code, 'Test14');
    // Empty character class should never match
    assert.strictEqual(result, vbaFalse, 'Empty [] should not match');
    console.log('[PASS] Test 14: Empty character class');
}

// Test 15: VBA veteran pattern - validating identifiers
{
    const code = `
    Function Test15()
        Dim isValid As Boolean
        isValid = ("myVar_1" Like "[a-zA-Z_]*")
        Test15 = isValid
    End Function
    `;
    const result = runFunc(code, 'Test15');
    assert.strictEqual(result, vbaTrue, 'Common pattern: validate identifier start');
    console.log('[PASS] Test 15: Identifier validation pattern');
}

console.log('\n✅ Like Operator Character Classes: 全テスト通過');
