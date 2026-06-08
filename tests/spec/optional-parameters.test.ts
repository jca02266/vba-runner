import { Evaluator } from '../../src/engine/evaluator';
import { evalVBASingle, assert, vbaTrue, vbaFalse } from '../../test-libs/test-runner';

function evalVBA(code: string): Evaluator {
    return evalVBASingle(code);
}

function runFunc(code: string, name: string, args: any[] = []): any {
    return evalVBA(code).callProcedure(name, args);
}

// Test 1: Optional parameter not provided - uses default value
{
    const code = String.raw`
        Function TestOptionalWithDefault(Optional x = 10)
            TestOptionalWithDefault = x
        End Function
    `;

    try {
        const result = runFunc(code, 'TestOptionalWithDefault');
        assert.strictEqual(result, 10, 'Optional parameter uses default value when not provided');
        console.log('[PASS] Optional parameter with default value');
    } catch (e: any) {
        console.log('[FAIL] Optional parameter with default value:', e.message);
    }
}

// Test 2: Optional parameter provided - uses provided value
{
    const code = String.raw`
        Function TestOptionalProvided(Optional x = 10)
            TestOptionalProvided = x
        End Function
    `;

    try {
        const result = runFunc(code, 'TestOptionalProvided', [42]);
        assert.strictEqual(result, 42, 'Optional parameter uses provided value');
        console.log('[PASS] Optional parameter with provided value');
    } catch (e: any) {
        console.log('[FAIL] Optional parameter with provided value:', e.message);
    }
}

// Test 3: IsMissing returns True when parameter not provided
{
    const code = String.raw`
        Function TestIsMissingTrue(Optional x)
            TestIsMissingTrue = IsMissing(x)
        End Function
    `;

    try {
        const result = runFunc(code, 'TestIsMissingTrue');
        assert.strictEqual(result, vbaTrue, 'IsMissing returns True when parameter not provided');
        console.log('[PASS] IsMissing returns True for missing parameter');
    } catch (e: any) {
        console.log('[FAIL] IsMissing returns True for missing parameter:', e.message);
    }
}

// Test 4: IsMissing returns False when parameter provided
{
    const code = String.raw`
        Function TestIsMissingFalse(Optional x)
            TestIsMissingFalse = IsMissing(x)
        End Function
    `;

    try {
        const result = runFunc(code, 'TestIsMissingFalse', [99]);
        assert.strictEqual(result, vbaFalse, 'IsMissing returns False when parameter provided');
        console.log('[PASS] IsMissing returns False for provided parameter');
    } catch (e: any) {
        console.log('[FAIL] IsMissing returns False for provided parameter:', e.message);
    }
}

// Test 5: Multiple optional parameters - some provided
{
    const code = String.raw`
        Function TestMultipleOptional(Optional a = 1, Optional b = 2, Optional c = 3)
            TestMultipleOptional = a + b + c
        End Function
    `;

    try {
        const result = runFunc(code, 'TestMultipleOptional', [10, 20]);
        // a=10, b=20, c=3 (default)
        assert.strictEqual(result, 33, 'Multiple optional parameters with partial arguments');
        console.log('[PASS] Multiple optional parameters');
    } catch (e: any) {
        console.log('[FAIL] Multiple optional parameters:', e.message);
    }
}

// Test 6: Optional parameter without default value
{
    const code = String.raw`
        Function TestOptionalNoDefault(Optional x)
            If IsMissing(x) Then
                TestOptionalNoDefault = -1
            Else
                TestOptionalNoDefault = x
            End If
        End Function
    `;

    try {
        const resultMissing = runFunc(code, 'TestOptionalNoDefault');
        assert.strictEqual(resultMissing, -1, 'Optional without default - missing returns -1');

        const resultProvided = runFunc(code, 'TestOptionalNoDefault', [99]);
        assert.strictEqual(resultProvided, 99, 'Optional without default - provided returns value');
        console.log('[PASS] Optional parameter without default value');
    } catch (e: any) {
        console.log('[FAIL] Optional parameter without default value:', e.message);
    }
}

// Test 7: Optional parameter with typed declaration
{
    const code = String.raw`
        Function TestOptionalTyped(Optional x As Integer = 5)
            TestOptionalTyped = x
        End Function
    `;

    try {
        const result = runFunc(code, 'TestOptionalTyped');
        assert.strictEqual(result, 5, 'Typed optional parameter uses default');
        console.log('[PASS] Typed optional parameter');
    } catch (e: any) {
        console.log('[FAIL] Typed optional parameter:', e.message);
    }
}

// Test 8: Optional Variant parameter with IsMissing
{
    const code = String.raw`
        Function TestOptionalVariant(Optional x As Variant)
            If IsMissing(x) Then
                TestOptionalVariant = 0
            Else
                TestOptionalVariant = x
            End If
        End Function
    `;

    try {
        const resultMissing = runFunc(code, 'TestOptionalVariant');
        assert.strictEqual(resultMissing, 0, 'Variant optional missing returns 0');

        const resultProvided = runFunc(code, 'TestOptionalVariant', [123]);
        assert.strictEqual(resultProvided, 123, 'Variant optional provided returns value');
        console.log('[PASS] Optional Variant parameter');
    } catch (e: any) {
        console.log('[FAIL] Optional Variant parameter:', e.message);
    }
}

// Test 9: Positional and optional parameters mixed
{
    const code = String.raw`
        Function TestMixedParameters(a, Optional b = 100, Optional c = 200)
            TestMixedParameters = a + b + c
        End Function
    `;

    try {
        const result = runFunc(code, 'TestMixedParameters', [10]);
        assert.strictEqual(result, 310, 'Mixed positional and optional parameters');
        console.log('[PASS] Mixed positional and optional parameters');
    } catch (e: any) {
        console.log('[FAIL] Mixed positional and optional parameters:', e.message);
    }
}

// Test 10: Real-world pattern - conditional behavior based on parameters
{
    const code = String.raw`
        Function FormatValue(value, Optional prefix, Optional suffix)
            Dim result
            result = value
            If Not IsMissing(prefix) Then
                result = prefix & result
            End If
            If Not IsMissing(suffix) Then
                result = result & suffix
            End If
            FormatValue = result
        End Function
    `;

    try {
        const resultNoPrefix = runFunc(code, 'FormatValue', ['test']);
        assert.strictEqual(resultNoPrefix, 'test', 'No prefix or suffix');

        const resultWithBoth = runFunc(code, 'FormatValue', ['test', '>', '<']);
        assert.strictEqual(resultWithBoth, '>test<', 'With prefix and suffix');

        const resultOnlyPrefix = runFunc(code, 'FormatValue', ['test', 'PREFIX:']);
        assert.strictEqual(resultOnlyPrefix, 'PREFIX:test', 'Only prefix');
        console.log('[PASS] Real-world conditional parameter handling');
    } catch (e: any) {
        console.log('[FAIL] Real-world conditional parameter handling:', e.message);
    }
}

// Test 11: Optional with string default value
{
    const code = String.raw`
        Function TestStringDefault(Optional msg = "default")
            TestStringDefault = msg
        End Function
    `;

    try {
        const resultDefault = runFunc(code, 'TestStringDefault');
        assert.strictEqual(resultDefault, 'default', 'String default value');

        const resultProvided = runFunc(code, 'TestStringDefault', ['custom']);
        assert.strictEqual(resultProvided, 'custom', 'String provided value');
        console.log('[PASS] Optional with string default value');
    } catch (e: any) {
        console.log('[FAIL] Optional with string default value:', e.message);
    }
}

// Test 12: IsMissing with non-optional parameter (should return False)
{
    const code = String.raw`
        Function TestIsMissingNonOptional(x)
            TestIsMissingNonOptional = IsMissing(x)
        End Function
    `;

    try {
        const result = runFunc(code, 'TestIsMissingNonOptional', [42]);
        assert.strictEqual(result, vbaFalse, 'IsMissing on non-optional parameter returns False');
        console.log('[PASS] IsMissing on non-optional parameter');
    } catch (e: any) {
        console.log('[FAIL] IsMissing on non-optional parameter:', e.message);
    }
}

console.log('\n✅ Optional Parameters: 全テスト完了');
