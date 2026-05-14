import { Lexer } from '../../src/compiler/lexer';
import { Parser } from '../../src/compiler/parser';
import { Evaluator } from '../../src/compiler/evaluator';
import { assert } from '../../test-libs/test-runner';

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

// Test 1: Basic StrConv with vbUpperCase
{
    const code = `
    Function Test1()
        Dim result As String
        result = StrConv("hello", vbUpperCase)
        Test1 = result
    End Function
    `;
    const result = runFunc(code, 'Test1');
    assert.strictEqual(result, 'HELLO', 'vbUpperCase should convert to uppercase');
    console.log('[PASS] Test 1: vbUpperCase');
}

// Test 2: Basic StrConv with vbLowerCase
{
    const code = `
    Function Test2()
        Dim result As String
        result = StrConv("HELLO", vbLowerCase)
        Test2 = result
    End Function
    `;
    const result = runFunc(code, 'Test2');
    assert.strictEqual(result, 'hello', 'vbLowerCase should convert to lowercase');
    console.log('[PASS] Test 2: vbLowerCase');
}

// Test 3: StrConv with vbProperCase
{
    const code = `
    Function Test3()
        Dim result As String
        result = StrConv("hello world", vbProperCase)
        Test3 = result
    End Function
    `;
    const result = runFunc(code, 'Test3');
    assert.strictEqual(result, 'Hello World', 'vbProperCase should capitalize first letter of each word');
    console.log('[PASS] Test 3: vbProperCase');
}

// Test 4: StrConv with combined flags (vbUpperCase + vbUnicode)
{
    const code = `
    Function Test4()
        Dim result As String
        result = StrConv("hello", vbUpperCase + vbUnicode)
        Test4 = result
    End Function
    `;
    const result = runFunc(code, 'Test4');
    assert.strictEqual(result, 'HELLO', 'vbUpperCase + vbUnicode should uppercase and convert to Unicode');
    console.log('[PASS] Test 4: Combined flags vbUpperCase + vbUnicode');
}

// Test 5: StrConv with vbUnicode flag
{
    const code = `
    Function Test5()
        Dim str As String
        Dim result As String
        str = "Hello"
        result = StrConv(str, vbUnicode)
        Test5 = result
    End Function
    `;
    const result = runFunc(code, 'Test5');
    // vbUnicode converts to Unicode - in JS, strings are already Unicode, so should be same
    assert.strictEqual(result, 'Hello', 'vbUnicode on ASCII string should be preserved');
    console.log('[PASS] Test 5: vbUnicode conversion');
}

// Test 6: StrConv with vbFromUnicode flag
{
    const code = `
    Function Test6()
        Dim str As String
        Dim result As String
        str = "Hello"
        result = StrConv(str, vbFromUnicode)
        Test6 = result
    End Function
    `;
    const result = runFunc(code, 'Test6');
    // vbFromUnicode converts from Unicode - in JS, strings are Unicode, should return same
    assert.strictEqual(result, 'Hello', 'vbFromUnicode on ASCII string should be preserved');
    console.log('[PASS] Test 6: vbFromUnicode conversion');
}

// Test 7: StrConv with vbLowerCase + vbUnicode
{
    const code = `
    Function Test7()
        Dim result As String
        result = StrConv("WORLD", vbLowerCase + vbUnicode)
        Test7 = result
    End Function
    `;
    const result = runFunc(code, 'Test7');
    assert.strictEqual(result, 'world', 'vbLowerCase + vbUnicode should lowercase and convert to Unicode');
    console.log('[PASS] Test 7: vbLowerCase + vbUnicode');
}

// Test 8: StrConv ProperCase with multiple words
{
    const code = `
    Function Test8()
        Dim result As String
        result = StrConv("the quick brown fox", vbProperCase)
        Test8 = result
    End Function
    `;
    const result = runFunc(code, 'Test8');
    assert.strictEqual(result, 'The Quick Brown Fox', 'vbProperCase should capitalize each word');
    console.log('[PASS] Test 8: vbProperCase multiple words');
}

// Test 9: StrConv with empty string
{
    const code = `
    Function Test9()
        Dim result As String
        result = StrConv("", vbUpperCase)
        Test9 = result
    End Function
    `;
    const result = runFunc(code, 'Test9');
    assert.strictEqual(result, '', 'StrConv on empty string should return empty');
    console.log('[PASS] Test 9: Empty string');
}

// Test 10: StrConv ProperCase with apostrophe
{
    const code = `
    Function Test10()
        Dim result As String
        result = StrConv("it's a test", vbProperCase)
        Test10 = result
    End Function
    `;
    const result = runFunc(code, 'Test10');
    // ProperCase treats apostrophe as part of the word (not a separator)
    // So "it's" becomes "It's", not "It'S"
    const expected = "It's A Test";
    assert.strictEqual(result, expected, 'vbProperCase treats apostrophe as part of word');
    console.log('[PASS] Test 10: ProperCase with apostrophe');
}

// Test 11: StrConv with numbers and special characters
{
    const code = `
    Function Test11()
        Dim result As String
        result = StrConv("hello123!@#", vbUpperCase)
        Test11 = result
    End Function
    `;
    const result = runFunc(code, 'Test11');
    assert.strictEqual(result, 'HELLO123!@#', 'StrConv should only affect letters');
    console.log('[PASS] Test 11: Numbers and special characters');
}

// Test 12: StrConv with Unicode literal
{
    const code = `
    Function Test12()
        Dim result As String
        result = StrConv("café", vbUpperCase)
        Test12 = result
    End Function
    `;
    const result = runFunc(code, 'Test12');
    // Unicode character should be handled
    assert.strictEqual(result.toUpperCase(), result, 'Unicode characters should be processed');
    console.log('[PASS] Test 12: Unicode literal');
}

// Test 13: StrConv vbProperCase with numbers
{
    const code = `
    Function Test13()
        Dim result As String
        result = StrConv("word2word test", vbProperCase)
        Test13 = result
    End Function
    `;
    const result = runFunc(code, 'Test13');
    assert.strictEqual(result, 'Word2word Test', 'vbProperCase should treat numbers as non-word boundaries');
    console.log('[PASS] Test 13: ProperCase with numbers');
}

// Test 14: StrConv multiple conversions in sequence
{
    const code = `
    Function Test14()
        Dim result As String
        result = "hello world"
        result = StrConv(result, vbProperCase)
        result = StrConv(result, vbUpperCase)
        Test14 = result
    End Function
    `;
    const result = runFunc(code, 'Test14');
    assert.strictEqual(result, 'HELLO WORLD', 'Sequential conversions should apply');
    console.log('[PASS] Test 14: Sequential conversions');
}

// Test 15: VBA veteran pattern - case insensitive comparison via StrConv
{
    const code = `
    Function Test15()
        Dim s As String
        Dim upper As String
        Dim lower As String
        s = "MixedCase"
        upper = StrConv(s, vbUpperCase)
        lower = StrConv(s, vbLowerCase)
        If (upper = "MIXEDCASE") And (lower = "mixedcase") Then
            Test15 = 1
        Else
            Test15 = 0
        End If
    End Function
    `;
    const result = runFunc(code, 'Test15');
    assert.strictEqual(result, 1, 'Common pattern: case conversion for comparison');
    console.log('[PASS] Test 15: VBA veteran pattern - case conversion');
}

console.log('\n✅ StrConv Unicode/ANSI Conversion: 全テスト通過');
