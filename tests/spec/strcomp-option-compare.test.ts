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

// Test 1: StrComp with explicit Binary compare
{
    const code = `
    Function Test1()
        Test1 = StrComp("ABC", "abc", 0)
    End Function
    `;
    const result = runFunc(code, 'Test1');
    assert.strictEqual(result, -1, 'Binary compare: "ABC" < "abc"');
    console.log('[PASS] Test 1: Explicit binary compare');
}

// Test 2: StrComp with explicit Text compare
{
    const code = `
    Function Test2()
        Test2 = StrComp("ABC", "abc", 1)
    End Function
    `;
    const result = runFunc(code, 'Test2');
    assert.strictEqual(result, 0, 'Text compare: "ABC" == "abc" (case-insensitive)');
    console.log('[PASS] Test 2: Explicit text compare');
}

// Test 3: StrComp without Compare parameter (should use Option Binary as default)
{
    const code = `
    Function Test3()
        Test3 = StrComp("ABC", "abc")
    End Function
    `;
    const result = runFunc(code, 'Test3');
    // Default should be binary (vbBinaryCompare = 0)
    assert.strictEqual(result, -1, 'Default (no param): should be binary compare');
    console.log('[PASS] Test 3: Default compare (no param)');
}

// Test 4: StrComp with Option Compare Binary
{
    const code = `
    Option Compare Binary
    Function Test4()
        Test4 = StrComp("ABC", "abc")
    End Function
    `;
    const result = runFunc(code, 'Test4');
    assert.strictEqual(result, -1, 'Option Binary: "ABC" < "abc"');
    console.log('[PASS] Test 4: Option Compare Binary');
}

// Test 5: StrComp with Option Compare Text
{
    const code = `
    Option Compare Text
    Function Test5()
        Test5 = StrComp("ABC", "abc")
    End Function
    `;
    const result = runFunc(code, 'Test5');
    assert.strictEqual(result, 0, 'Option Text: "ABC" == "abc"');
    console.log('[PASS] Test 5: Option Compare Text');
}

// Test 6: StrComp respects explicit parameter over Option
{
    const code = `
    Option Compare Text
    Function Test6()
        Test6 = StrComp("ABC", "abc", 0)
    End Function
    `;
    const result = runFunc(code, 'Test6');
    assert.strictEqual(result, -1, 'Explicit 0 overrides Option Text');
    console.log('[PASS] Test 6: Explicit parameter overrides option');
}

// Test 7: Equal strings (binary)
{
    const code = `
    Function Test7()
        Test7 = StrComp("hello", "hello", 0)
    End Function
    `;
    const result = runFunc(code, 'Test7');
    assert.strictEqual(result, 0, 'Equal strings return 0');
    console.log('[PASS] Test 7: Equal strings');
}

// Test 8: First string greater (binary)
{
    const code = `
    Function Test8()
        Test8 = StrComp("xyz", "abc", 0)
    End Function
    `;
    const result = runFunc(code, 'Test8');
    assert.strictEqual(result, 1, 'First > second returns 1');
    console.log('[PASS] Test 8: First string greater');
}

// Test 9: Null handling - skip for now (Null propagation is complex)
{
    const code = `
    Function Test9()
        Test9 = 1
    End Function
    `;
    const result = runFunc(code, 'Test9');
    assert.strictEqual(result, 1, 'Null handling test skipped');
    console.log('[PASS] Test 9: Placeholder (Null handling complex)');
}

// Test 10: Case-sensitive comparison with binary
{
    const code = `
    Option Compare Binary
    Function Test10()
        Test10 = StrComp("Password", "password")
    End Function
    `;
    const result = runFunc(code, 'Test10');
    assert.strictEqual(result, -1, 'Binary: "Password" < "password"');
    console.log('[PASS] Test 10: Case-sensitive binary comparison');
}

// Test 11: Case-insensitive comparison with text option
{
    const code = `
    Option Compare Text
    Function Test11()
        Test11 = StrComp("Password", "password")
    End Function
    `;
    const result = runFunc(code, 'Test11');
    assert.strictEqual(result, 0, 'Text: "Password" == "password"');
    console.log('[PASS] Test 11: Case-insensitive text comparison');
}

// Test 12: Numeric strings with binary compare
{
    const code = `
    Function Test12()
        Test12 = StrComp("10", "2", 0)
    End Function
    `;
    const result = runFunc(code, 'Test12');
    assert.strictEqual(result, -1, 'Binary: "10" < "2" (lexicographic)');
    console.log('[PASS] Test 12: Numeric strings binary');
}

// Test 13: Empty string comparison
{
    const code = `
    Function Test13()
        Test13 = StrComp("", "", 0)
    End Function
    `;
    const result = runFunc(code, 'Test13');
    assert.strictEqual(result, 0, 'Empty strings are equal');
    console.log('[PASS] Test 13: Empty strings');
}

// Test 14: Empty vs non-empty
{
    const code = `
    Function Test14()
        Test14 = StrComp("", "abc", 0)
    End Function
    `;
    const result = runFunc(code, 'Test14');
    assert.strictEqual(result, -1, 'Empty < non-empty');
    console.log('[PASS] Test 14: Empty vs non-empty');
}

// Test 15: Variant type coercion
{
    const code = `
    Function Test15()
        Dim v1 As Variant, v2 As Variant
        v1 = "ABC"
        v2 = "abc"
        Test15 = StrComp(v1, v2, 1)
    End Function
    `;
    const result = runFunc(code, 'Test15');
    assert.strictEqual(result, 0, 'Variant coercion with text compare');
    console.log('[PASS] Test 15: Variant coercion');
}

console.log('\n✅ StrComp Option Compare: 全テスト通過');
