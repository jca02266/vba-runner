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

// Test 1: Variant + Byte → Variant type
{
    const code = `
        Function TestVariantPlusByte()
            Dim v As Variant
            v = 10
            Dim b As Byte
            b = 5
            TestVariantPlusByte = v + b  ' Should be Variant type with value 15
        End Function
    `;
    const result = runFunc(code, 'TestVariantPlusByte');
    assert.strictEqual(result, 15, 'Variant + Byte arithmetic');
    console.log('[PASS] Variant + Byte → result is 15');
}

// Test 2: Variant + Integer → Variant type
{
    const code = `
        Function TestVariantPlusInteger()
            Dim v As Variant
            v = 100
            Dim i As Integer
            i = 50
            TestVariantPlusInteger = v + i
        End Function
    `;
    const result = runFunc(code, 'TestVariantPlusInteger');
    assert.strictEqual(result, 150, 'Variant + Integer arithmetic');
    console.log('[PASS] Variant + Integer → result is 150');
}

// Test 3: Variant + Double → Variant type
{
    const code = `
        Function TestVariantPlusDouble()
            Dim v As Variant
            v = 10.5
            Dim d As Double
            d = 5.5
            TestVariantPlusDouble = v + d
        End Function
    `;
    const result = runFunc(code, 'TestVariantPlusDouble');
    assert.strictEqual(result, 16, 'Variant + Double arithmetic');
    console.log('[PASS] Variant + Double → result is 16');
}

// Test 4: Variant + String (numeric string) → Variant type
{
    const code = `
        Function TestVariantPlusNumericString()
            Dim v As Variant
            v = 10
            Dim s As String
            s = "5"
            TestVariantPlusNumericString = v + s  ' String coerced to 5
        End Function
    `;
    const result = runFunc(code, 'TestVariantPlusNumericString');
    assert.strictEqual(result, 15, 'Variant + numeric string coercion');
    console.log('[PASS] Variant + numeric string → result is 15');
}

// Test 5: Unary Minus on Variant
{
    const code = `
        Function TestVariantUnaryMinus()
            Dim v As Variant
            v = 42
            TestVariantUnaryMinus = -v  ' Should be -42
        End Function
    `;
    const result = runFunc(code, 'TestVariantUnaryMinus');
    assert.strictEqual(result, -42, 'Unary minus on Variant');
    console.log('[PASS] Unary - Variant → result is -42');
}

// Test 6: Variant - Integer (subtraction)
{
    const code = `
        Function TestVariantMinus()
            Dim v As Variant
            v = 100
            TestVariantMinus = v - 30
        End Function
    `;
    const result = runFunc(code, 'TestVariantMinus');
    assert.strictEqual(result, 70, 'Variant - integer subtraction');
    console.log('[PASS] Variant - Integer → result is 70');
}

// Test 7: Variant * Integer (multiplication)
{
    const code = `
        Function TestVariantMultiply()
            Dim v As Variant
            v = 12
            TestVariantMultiply = v * 3
        End Function
    `;
    const result = runFunc(code, 'TestVariantMultiply');
    assert.strictEqual(result, 36, 'Variant * integer multiplication');
    console.log('[PASS] Variant * Integer → result is 36');
}

// Test 8: Variant / Integer (division)
{
    const code = `
        Function TestVariantDivide()
            Dim v As Variant
            v = 100
            TestVariantDivide = v / 4
        End Function
    `;
    const result = runFunc(code, 'TestVariantDivide');
    assert.strictEqual(result, 25, 'Variant / integer division');
    console.log('[PASS] Variant / Integer → result is 25');
}

// Test 9: Integer + Variant (commutative)
{
    const code = `
        Function TestIntegerPlusVariant()
            Dim v As Variant
            v = 20
            TestIntegerPlusVariant = 30 + v  ' Order reversed
        End Function
    `;
    const result = runFunc(code, 'TestIntegerPlusVariant');
    assert.strictEqual(result, 50, 'Integer + Variant (commutative)');
    console.log('[PASS] Integer + Variant → result is 50');
}

// Test 10: Variant with Empty (vbaEmpty)
{
    const code = `
        Function TestVariantEmpty()
            Dim v As Variant
            ' v starts as Empty
            Dim result
            result = v + 10  ' Empty coerces to 0
            TestVariantEmpty = result
        End Function
    `;
    const result = runFunc(code, 'TestVariantEmpty');
    assert.strictEqual(result, 10, 'Empty in Variant coerces to 0');
    console.log('[PASS] Variant(Empty) + Integer → result is 10');
}

// Test 11: Chained operations with Variant
{
    const code = `
        Function TestVariantChained()
            Dim v As Variant
            v = 10
            TestVariantChained = v + 5 + 3  ' (10 + 5) + 3 = 18
        End Function
    `;
    const result = runFunc(code, 'TestVariantChained');
    assert.strictEqual(result, 18, 'Chained Variant operations');
    console.log('[PASS] Variant chained + operations → result is 18');
}

// Test 12: Variant mod Integer
{
    const code = `
        Function TestVariantMod()
            Dim v As Variant
            v = 17
            TestVariantMod = v Mod 5  ' 17 mod 5 = 2
        End Function
    `;
    const result = runFunc(code, 'TestVariantMod');
    assert.strictEqual(result, 2, 'Variant Mod integer');
    console.log('[PASS] Variant Mod Integer → result is 2');
}

// Test 13: Variant ^ Integer (exponentiation)
{
    const code = `
        Function TestVariantPower()
            Dim v As Variant
            v = 2
            TestVariantPower = v ^ 8  ' 2^8 = 256
        End Function
    `;
    const result = runFunc(code, 'TestVariantPower');
    assert.strictEqual(result, 256, 'Variant exponentiation');
    console.log('[PASS] Variant ^ Integer → result is 256');
}

// Test 14: Variant integer division (\)
{
    const code = `
        Function TestVariantIntDivide()
            Dim v As Variant
            v = 25
            TestVariantIntDivide = v \\ 4  ' 25 \ 4 = 6
        End Function
    `;
    const result = runFunc(code, 'TestVariantIntDivide');
    assert.strictEqual(result, 6, 'Variant integer division');
    console.log('[PASS] Variant \\ Integer → result is 6');
}

// Test 15: Real-world pattern - Dynamic calculation
{
    const code = `
        Function CalculateTotal(unitPrice, quantity)
            ' Both parameters are Variant (no type specified)
            ' This demonstrates the practical use of Variant propagation
            Dim subtotal
            subtotal = unitPrice * quantity  ' Variant * Variant
            Dim tax
            tax = subtotal * 0.1  ' Variant * Double
            CalculateTotal = subtotal + tax  ' Variant + Variant
        End Function

        Function TestDynamicCalculation()
            TestDynamicCalculation = CalculateTotal(29.99, 3)
        End Function
    `;
    const result = runFunc(code, 'TestDynamicCalculation');
    // 29.99 * 3 = 89.97, tax = 8.997, total ≈ 98.967
    // Allow small floating point variance
    assert.strictEqual(Math.abs(result - 98.967) < 0.01, true, 'Real-world variant calculation');
    console.log('[PASS] Real-world dynamic calculation with Variants');
}

console.log('\n✅ Variant Propagation: 全テスト通過');
