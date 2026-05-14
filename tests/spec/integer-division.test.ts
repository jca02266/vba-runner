import { Lexer } from '../../src/compiler/lexer';
import { Parser } from '../../src/compiler/parser';
import { Evaluator } from '../../src/compiler/evaluator';
import { assert } from '../../test-libs/test-runner';

function evalVBA(code: string): Evaluator {
    const tokens = new Lexer(code).tokenize();
    const ast = new Parser(tokens).parse();
    const ev = new Evaluator(console.log);
    ev.evaluate(ast);
    return ev;
}

function runFunc(code: string, name: string, args: any[] = []): any {
    return evalVBA(code).callProcedure(name, args);
}

// Test 1: Basic integer division - positive integers
{
    const code = String.raw`
        Function TestBasicIntDiv()
            Dim result
            result = 10 \ 3
            TestBasicIntDiv = result
        End Function
    `;

    try {
        const result = runFunc(code, 'TestBasicIntDiv');
        assert.strictEqual(result, 3, 'Basic integer division 10 \\ 3 = 3');
        console.log('[PASS] Basic integer division - positive');
    } catch (e: any) {
        console.log('[FAIL] Basic integer division - positive:', e.message);
    }
}

// Test 2: Integer division with exact result
{
    const code = String.raw`
        Function TestExactIntDiv()
            Dim result
            result = 10 \ 2
            TestExactIntDiv = result
        End Function
    `;

    try {
        const result = runFunc(code, 'TestExactIntDiv');
        assert.strictEqual(result, 5, 'Integer division with exact result');
        console.log('[PASS] Integer division with exact result');
    } catch (e: any) {
        console.log('[FAIL] Integer division with exact result:', e.message);
    }
}

// Test 3: Integer division - rounding toward zero (positive dividend)
{
    const code = String.raw`
        Function TestRoundTowardZeroPos()
            Dim result
            result = 7 \ 2
            TestRoundTowardZeroPos = result
        End Function
    `;

    try {
        const result = runFunc(code, 'TestRoundTowardZeroPos');
        assert.strictEqual(result, 3, 'Rounding toward zero: 7 \\ 2 = 3');
        console.log('[PASS] Integer division - rounding toward zero (positive)');
    } catch (e: any) {
        console.log('[FAIL] Integer division - rounding toward zero (positive):', e.message);
    }
}

// Test 4: Integer division - negative dividend (rounding toward zero)
{
    const code = String.raw`
        Function TestRoundTowardZeroNeg()
            Dim result
            result = -7 \ 2
            TestRoundTowardZeroNeg = result
        End Function
    `;

    try {
        const result = runFunc(code, 'TestRoundTowardZeroNeg');
        assert.strictEqual(result, -3, 'Rounding toward zero: -7 \\ 2 = -3 (not -4)');
        console.log('[PASS] Integer division - rounding toward zero (negative dividend)');
    } catch (e: any) {
        console.log('[FAIL] Integer division - rounding toward zero (negative dividend):', e.message);
    }
}

// Test 5: Integer division - negative divisor (rounding toward zero)
{
    const code = String.raw`
        Function TestRoundTowardZeroNegDiv()
            Dim result
            result = 7 \ -2
            TestRoundTowardZeroNegDiv = result
        End Function
    `;

    try {
        const result = runFunc(code, 'TestRoundTowardZeroNegDiv');
        assert.strictEqual(result, -3, 'Rounding toward zero: 7 \\ -2 = -3');
        console.log('[PASS] Integer division - rounding toward zero (negative divisor)');
    } catch (e: any) {
        console.log('[FAIL] Integer division - rounding toward zero (negative divisor):', e.message);
    }
}

// Test 6: Integer division - both negative (rounding toward zero)
{
    const code = String.raw`
        Function TestRoundTowardZeroBothNeg()
            Dim result
            result = -7 \ -2
            TestRoundTowardZeroBothNeg = result
        End Function
    `;

    try {
        const result = runFunc(code, 'TestRoundTowardZeroBothNeg');
        assert.strictEqual(result, 3, 'Rounding toward zero: -7 \\ -2 = 3');
        console.log('[PASS] Integer division - rounding toward zero (both negative)');
    } catch (e: any) {
        console.log('[FAIL] Integer division - rounding toward zero (both negative):', e.message);
    }
}

// Test 7: Integer division with Single type (automatic conversion to Long)
{
    const code = String.raw`
        Function TestSingleToIntDiv()
            Dim x As Single
            Dim y As Single
            Dim result
            x = 10.7
            y = 3.2
            result = x \ y
            TestSingleToIntDiv = result
        End Function
    `;

    try {
        const result = runFunc(code, 'TestSingleToIntDiv');
        assert.strictEqual(result, 3, 'Single converted to Long: 10 \\ 3 = 3');
        console.log('[PASS] Integer division with Single type');
    } catch (e: any) {
        console.log('[FAIL] Integer division with Single type:', e.message);
    }
}

// Test 8: Integer division with Double type (automatic conversion)
{
    const code = String.raw`
        Function TestDoubleToIntDiv()
            Dim x As Double
            Dim y As Double
            Dim result
            x = 25.9
            y = 4.1
            result = x \ y
            TestDoubleToIntDiv = result
        End Function
    `;

    try {
        const result = runFunc(code, 'TestDoubleToIntDiv');
        assert.strictEqual(result, 6, 'Double converted to Long: 25 \\ 4 = 6');
        console.log('[PASS] Integer division with Double type');
    } catch (e: any) {
        console.log('[FAIL] Integer division with Double type:', e.message);
    }
}

// Test 9: Integer division with String type (automatic conversion)
{
    const code = String.raw`
        Function TestStringToIntDiv()
            Dim x
            Dim y
            Dim result
            x = "15"
            y = "4"
            result = x \ y
            TestStringToIntDiv = result
        End Function
    `;

    try {
        const result = runFunc(code, 'TestStringToIntDiv');
        assert.strictEqual(result, 3, 'String converted to numeric: "15" \\ "4" = 3');
        console.log('[PASS] Integer division with String type');
    } catch (e: any) {
        console.log('[FAIL] Integer division with String type:', e.message);
    }
}

// Test 10: Integer division by zero - Error 11
{
    const code = String.raw`
        Function TestDivisionByZero()
            Dim result
            result = 10 \ 0
            TestDivisionByZero = result
        End Function
    `;

    try {
        runFunc(code, 'TestDivisionByZero');
        console.log('[FAIL] Division by zero not detected');
    } catch (e: any) {
        if (e && e.type === 'VbaError' && e.number === 11) {
            console.log('[PASS] Division by zero detected (Error 11)');
        } else {
            console.log('[FAIL] Wrong error type:', e.message);
        }
    }
}

// Test 11: Integer division with Byte type
{
    const code = String.raw`
        Function TestByteIntDiv()
            Dim x As Byte
            Dim y As Byte
            Dim result
            x = 200
            y = 7
            result = x \ y
            TestByteIntDiv = result
        End Function
    `;

    try {
        const result = runFunc(code, 'TestByteIntDiv');
        assert.strictEqual(result, 28, 'Byte division: 200 \\ 7 = 28');
        console.log('[PASS] Integer division with Byte type');
    } catch (e: any) {
        console.log('[FAIL] Integer division with Byte type:', e.message);
    }
}

// Test 12: Integer division with Integer type
{
    const code = String.raw`
        Function TestIntegerIntDiv()
            Dim x As Integer
            Dim y As Integer
            Dim result
            x = 1000
            y = 3
            result = x \ y
            TestIntegerIntDiv = result
        End Function
    `;

    try {
        const result = runFunc(code, 'TestIntegerIntDiv');
        assert.strictEqual(result, 333, 'Integer division: 1000 \\ 3 = 333');
        console.log('[PASS] Integer division with Integer type');
    } catch (e: any) {
        console.log('[FAIL] Integer division with Integer type:', e.message);
    }
}

// Test 13: Integer division with Long type
{
    const code = String.raw`
        Function TestLongIntDiv()
            Dim x As Long
            Dim y As Long
            Dim result
            x = 2000000
            y = 7
            result = x \ y
            TestLongIntDiv = result
        End Function
    `;

    try {
        const result = runFunc(code, 'TestLongIntDiv');
        assert.strictEqual(result, 285714, 'Long division: 2000000 \\ 7 = 285714');
        console.log('[PASS] Integer division with Long type');
    } catch (e: any) {
        console.log('[FAIL] Integer division with Long type:', e.message);
    }
}

// Test 14: Mod operator - basic positive
{
    const code = String.raw`
        Function TestBasicMod()
            Dim result
            result = 10 Mod 3
            TestBasicMod = result
        End Function
    `;

    try {
        const result = runFunc(code, 'TestBasicMod');
        assert.strictEqual(result, 1, 'Basic modulo: 10 Mod 3 = 1');
        console.log('[PASS] Basic Mod operator - positive');
    } catch (e: any) {
        console.log('[FAIL] Basic Mod operator - positive:', e.message);
    }
}

// Test 15: Mod operator - with negative dividend
{
    const code = String.raw`
        Function TestModNegDividend()
            Dim result
            result = -10 Mod 3
            TestModNegDividend = result
        End Function
    `;

    try {
        const result = runFunc(code, 'TestModNegDividend');
        // In VBA, -10 Mod 3 = -1 (sign follows dividend)
        assert.strictEqual(result, -1, 'Mod with negative dividend: -10 Mod 3 = -1');
        console.log('[PASS] Mod operator - negative dividend');
    } catch (e: any) {
        console.log('[FAIL] Mod operator - negative dividend:', e.message);
    }
}

// Test 16: Real-world pattern - calculations using integer division
{
    const code = String.raw`
        Function TestCalculationWithIntDiv()
            Dim hours
            Dim minutes
            Dim totalSeconds
            totalSeconds = 3661
            hours = totalSeconds \ 3600
            minutes = (totalSeconds Mod 3600) \ 60
            TestCalculationWithIntDiv = hours * 60 + minutes
        End Function
    `;

    try {
        const result = runFunc(code, 'TestCalculationWithIntDiv');
        assert.strictEqual(result, 61, 'Time calculation: 1h 1m = 61m');
        console.log('[PASS] Real-world calculation with integer division');
    } catch (e: any) {
        console.log('[FAIL] Real-world calculation with integer division:', e.message);
    }
}

// Test 17: Chained integer division
{
    const code = String.raw`
        Function TestChainedIntDiv()
            Dim result
            result = 100 \ 3 \ 2
            TestChainedIntDiv = result
        End Function
    `;

    try {
        const result = runFunc(code, 'TestChainedIntDiv');
        assert.strictEqual(result, 16, 'Chained division: (100 \\ 3) \\ 2 = 16');
        console.log('[PASS] Chained integer division');
    } catch (e: any) {
        console.log('[FAIL] Chained integer division:', e.message);
    }
}

// Test 18: Integer division mixed with other operators
{
    const code = String.raw`
        Function TestMixedOps()
            Dim result
            result = 10 \ 3 + 5 \ 2
            TestMixedOps = result
        End Function
    `;

    try {
        const result = runFunc(code, 'TestMixedOps');
        assert.strictEqual(result, 5, 'Mixed operators: (10 \\ 3) + (5 \\ 2) = 5');
        console.log('[PASS] Integer division mixed with other operators');
    } catch (e: any) {
        console.log('[FAIL] Integer division mixed with other operators:', e.message);
    }
}

// Test 19: Integer division with variable assignment
{
    const code = String.raw`
        Function TestWithVariables()
            Dim x
            Dim y
            Dim z
            x = 100
            y = 7
            z = x \ y
            TestWithVariables = z
        End Function
    `;

    try {
        const result = runFunc(code, 'TestWithVariables');
        assert.strictEqual(result, 14, 'Variable division: 100 \\ 7 = 14');
        console.log('[PASS] Integer division with variable assignment');
    } catch (e: any) {
        console.log('[FAIL] Integer division with variable assignment:', e.message);
    }
}

// Test 20: Mod operator with negative divisor
{
    const code = String.raw`
        Function TestModNegDivisor()
            Dim result
            result = 10 Mod -3
            TestModNegDivisor = result
        End Function
    `;

    try {
        const result = runFunc(code, 'TestModNegDivisor');
        // In VBA, 10 Mod -3 = 1 (sign follows divisor)
        assert.strictEqual(result, 1, 'Mod with negative divisor: 10 Mod -3 = 1');
        console.log('[PASS] Mod operator - negative divisor');
    } catch (e: any) {
        console.log('[FAIL] Mod operator - negative divisor:', e.message);
    }
}

console.log('\n✅ Integer Division: 全テスト完了');
