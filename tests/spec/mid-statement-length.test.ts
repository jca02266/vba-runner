import { Lexer } from '../../src/engine/lexer';
import { Parser } from '../../src/engine/parser';
import { Evaluator } from '../../src/engine/evaluator';
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

// Test 1: Basic Mid$ with length parameter - replacement shorter than allocated space
{
    const code = `
    Function Test1()
        Dim s As String
        s = "Hello World"
        Mid$(s, 1, 5) = "HI"
        Test1 = s
    End Function
    `;
    const result = runFunc(code, 'Test1');
    assert.strictEqual(result, 'HIllo World', 'Mid$(s, 1, 5) = "HI" replaces min(5, 11, 2)=2 chars with "HI"');
    console.log('[PASS] Test 1: Basic Mid$ with length - replacement shorter than allocated');
}

// Test 2: Mid$ with length parameter - replacement exactly fits
{
    const code = `
    Function Test2()
        Dim s As String
        s = "Hello"
        Mid$(s, 1, 5) = "WORLD"
        Test2 = s
    End Function
    `;
    const result = runFunc(code, 'Test2');
    assert.strictEqual(result, 'WORLD', 'Mid$(s, 1, 5) = "WORLD" should replace entire string');
    console.log('[PASS] Test 2: Mid$ with length - exact fit');
}

// Test 3: Mid$ with length parameter - replacement longer than allocated space
{
    const code = `
    Function Test3()
        Dim s As String
        s = "Hello"
        Mid$(s, 1, 3) = "WXYZABC"
        Test3 = s
    End Function
    `;
    const result = runFunc(code, 'Test3');
    assert.strictEqual(result, 'WXYlo', 'Mid$(s, 1, 3) = "WXYZABC" should only use first 3 chars of replacement');
    console.log('[PASS] Test 3: Mid$ with length - replacement longer than allocated');
}

// Test 4: Mid$ starting mid-string with length parameter
{
    const code = `
    Function Test4()
        Dim s As String
        s = "Hello World"
        Mid$(s, 7, 3) = "VB"
        Test4 = s
    End Function
    `;
    const result = runFunc(code, 'Test4');
    assert.strictEqual(result, 'Hello VBrld', 'Mid$(s, 7, 3) = "VB" replaces chars 7-9');
    console.log('[PASS] Test 4: Mid$ starting mid-string with length');
}

// Test 5: Mid$ without length parameter - replacement shorter than remaining
{
    const code = `
    Function Test5()
        Dim s As String
        s = "Hello World"
        Mid$(s, 7) = "VB"
        Test5 = s
    End Function
    `;
    const result = runFunc(code, 'Test5');
    assert.strictEqual(result, 'Hello VBrld', 'Mid$(s, 7) = "VB" without length replaces to end or fills replacement');
    console.log('[PASS] Test 5: Mid$ without length - replacement shorter');
}

// Test 6: Mid$ without length parameter - replacement longer than remaining
{
    const code = `
    Function Test6()
        Dim s As String
        s = "Hello"
        Mid$(s, 2) = "WXYZABC"
        Test6 = s
    End Function
    `;
    const result = runFunc(code, 'Test6');
    assert.strictEqual(result, 'HWXYZ', 'Mid$(s, 2) = "WXYZABC" replaces min(no-limit, 4-remaining, 7)=4 chars');
    console.log('[PASS] Test 6: Mid$ without length - replacement longer');
}

// Test 7: Mid$ at end of string with length
{
    const code = `
    Function Test7()
        Dim s As String
        s = "Hello"
        Mid$(s, 5, 10) = "XY"
        Test7 = s
    End Function
    `;
    const result = runFunc(code, 'Test7');
    assert.strictEqual(result, 'HellX', 'Mid$(s, 5, 10) = "XY" replaces min(10, 1, 2)=1 char at end');
    console.log('[PASS] Test 7: Mid$ at end of string');
}

// Test 8: Mid$ replacing single character
{
    const code = `
    Function Test8()
        Dim s As String
        s = "Hello"
        Mid$(s, 2, 1) = "A"
        Test8 = s
    End Function
    `;
    const result = runFunc(code, 'Test8');
    assert.strictEqual(result, 'HAllo', 'Mid$(s, 2, 1) = "A" replaces character at position 2');
    console.log('[PASS] Test 8: Single character replacement');
}

// Test 9: Mid$ with empty replacement string
{
    const code = `
    Function Test9()
        Dim s As String
        s = "Hello"
        Mid$(s, 2, 3) = ""
        Test9 = s
    End Function
    `;
    const result = runFunc(code, 'Test9');
    assert.strictEqual(result, 'Hello', 'Mid$(s, 2, 3) = "" replaces min(3, 4, 0)=0 chars, so no change');
    console.log('[PASS] Test 9: Empty replacement string');
}

// Test 10: Mid$ with various string types (variant)
{
    const code = `
    Function Test10()
        Dim s As Variant
        s = "Hello"
        Mid$(s, 1, 2) = "HI"
        Test10 = s
    End Function
    `;
    const result = runFunc(code, 'Test10');
    assert.strictEqual(result, 'HIllo', 'Mid$ should work with Variant type');
    console.log('[PASS] Test 10: Variant type support');
}

// Test 11: Mid$ replacement preserves rest of string
{
    const code = `
    Function Test11()
        Dim s As String
        s = "The quick brown fox"
        Mid$(s, 5, 5) = "SLOW"
        Test11 = s
    End Function
    `;
    const result = runFunc(code, 'Test11');
    assert.strictEqual(result, 'The SLOWk brown fox', 'Mid$(s, 5, 5) = "SLOW" replaces min(5,15,4)=4 chars');
    console.log('[PASS] Test 11: Rest of string preservation');
}

// Test 12: Mid$ multiple operations
{
    const code = `
    Function Test12()
        Dim s As String
        s = "ABCDEF"
        Mid$(s, 1, 2) = "12"
        Mid$(s, 4, 2) = "45"
        Test12 = s
    End Function
    `;
    const result = runFunc(code, 'Test12');
    assert.strictEqual(result, '12C45F', 'Multiple Mid$ operations should accumulate');
    console.log('[PASS] Test 12: Multiple operations');
}

// Test 13: Mid$ with start position at string boundary
{
    const code = `
    Function Test13()
        Dim s As String
        s = "Hi"
        Mid$(s, 2, 5) = "THERE"
        Test13 = s
    End Function
    `;
    const result = runFunc(code, 'Test13');
    assert.strictEqual(result, 'HT', 'Mid$(s, 2, 5) = "THERE" replaces min(5, 1, 5)=1 char at position 2');
    console.log('[PASS] Test 13: Start at boundary');
}

// Test 14: Mid$ with length greater than remaining but shorter replacement
{
    const code = `
    Function Test14()
        Dim s As String
        s = "ABC"
        Mid$(s, 2, 10) = "X"
        Test14 = s
    End Function
    `;
    const result = runFunc(code, 'Test14');
    assert.strictEqual(result, 'AXC', 'Should respect actual remaining chars even if length says more');
    console.log('[PASS] Test 14: Length > remaining, but short replacement');
}

// Test 15: VBA veteran pattern - building strings with Mid$
{
    const code = `
    Function Test15()
        Dim result As String
        result = "____"
        Mid$(result, 1, 1) = "H"
        Mid$(result, 2, 1) = "i"
        Test15 = result
    End Function
    `;
    const result = runFunc(code, 'Test15');
    assert.strictEqual(result, 'Hi__', 'Common pattern: pre-allocate string and fill with Mid$');
    console.log('[PASS] Test 15: Pre-allocate pattern');
}

console.log('\n✅ Mid$ Statement Length Rules: 全テスト通過');
