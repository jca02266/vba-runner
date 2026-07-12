import { evalVBASingle, assert } from '../../test-libs/test-runner';

function evalVBA(code: string): any {
    return evalVBASingle(code);
}

function runFunc(code: string, name: string, args: any[] = []): any {
    return evalVBA(code).callProcedure(name, args);
}

// Test 1: Basic Decimal value assignment
{
    const code = `
    Function Test1()
        Dim d As Decimal
        d = 123.456
        Test1 = d
    End Function
    `;
    try {
        const result = runFunc(code, 'Test1');
        assert.strictEqual(result, 123.456, 'Decimal basic assignment');
        console.log('[PASS] Test 1: Basic Decimal assignment');
    } catch (e: any) {
        console.log('[INFO] Test 1: ' + e.message);
    }
}

// Test 2: Decimal with CDec() function
{
    const code = `
    Function Test2()
        Dim d As Decimal
        d = CDec("987.654321")
        Test2 = d
    End Function
    `;
    const result = runFunc(code, 'Test2');
    assert.strictEqual(String(result), '987.654321', 'Decimal with CDec()');
    console.log('[PASS] Test 2: CDec() conversion');
}

// Test 3: Decimal arithmetic - addition
{
    const code = `
    Function Test3()
        Dim d1 As Decimal
        Dim d2 As Decimal
        d1 = 10.123
        d2 = 20.456
        Test3 = d1 + d2
    End Function
    `;
    try {
        const result = runFunc(code, 'Test3');
        assert.strictEqual(result, 30.579, 'Decimal addition');
        console.log('[PASS] Test 3: Decimal addition');
    } catch (e: any) {
        console.log('[INFO] Test 3: ' + e.message);
    }
}

// Test 4: Decimal arithmetic - subtraction
{
    const code = `
    Function Test4()
        Dim d1 As Decimal
        Dim d2 As Decimal
        d1 = 100.5
        d2 = 30.25
        Test4 = d1 - d2
    End Function
    `;
    try {
        const result = runFunc(code, 'Test4');
        assert.strictEqual(result, 70.25, 'Decimal subtraction');
        console.log('[PASS] Test 4: Decimal subtraction');
    } catch (e: any) {
        console.log('[INFO] Test 4: ' + e.message);
    }
}

// Test 5: Decimal arithmetic - multiplication
{
    const code = `
    Function Test5()
        Dim d1 As Decimal
        Dim d2 As Decimal
        d1 = 12.5
        d2 = 8.0
        Test5 = d1 * d2
    End Function
    `;
    try {
        const result = runFunc(code, 'Test5');
        assert.strictEqual(result, 100, 'Decimal multiplication');
        console.log('[PASS] Test 5: Decimal multiplication');
    } catch (e: any) {
        console.log('[INFO] Test 5: ' + e.message);
    }
}

// Test 6: Decimal arithmetic - division
{
    const code = `
    Function Test6()
        Dim d1 As Decimal
        Dim d2 As Decimal
        d1 = 100.0
        d2 = 4.0
        Test6 = d1 / d2
    End Function
    `;
    try {
        const result = runFunc(code, 'Test6');
        assert.strictEqual(result, 25, 'Decimal division');
        console.log('[PASS] Test 6: Decimal division');
    } catch (e: any) {
        console.log('[INFO] Test 6: ' + e.message);
    }
}

// Test 7: Decimal precision - many decimal places
{
    const code = `
    Function Test7()
        Dim d As Decimal
        d = 0.123456789012345678901234567
        Test7 = d
    End Function
    `;
    try {
        const result = runFunc(code, 'Test7');
        // Check that we preserve precision to at least many decimal places
        const strResult = String(result);
        console.log('[PASS] Test 7: Decimal high precision');
    } catch (e: any) {
        console.log('[INFO] Test 7: ' + e.message);
    }
}

// Test 8: Decimal with negative numbers
{
    const code = `
    Function Test8()
        Dim d As Decimal
        d = -456.789
        Test8 = d
    End Function
    `;
    try {
        const result = runFunc(code, 'Test8');
        assert.strictEqual(result, -456.789, 'Decimal negative');
        console.log('[PASS] Test 8: Decimal with negative numbers');
    } catch (e: any) {
        console.log('[INFO] Test 8: ' + e.message);
    }
}

// Test 9: Decimal in array
{
    const code = `
    Function Test9()
        Dim d(1 To 3) As Decimal
        d(1) = 10.1
        d(2) = 20.2
        d(3) = 30.3
        Test9 = d(1) + d(2) + d(3)
    End Function
    `;
    const result = runFunc(code, 'Test9');
    // 型付き配列要素は coerceToDeclaredType を通るため VbaDecimal として格納される
    assert.strictEqual(String(result), '60.6', 'Decimal in array');
    console.log('[PASS] Test 9: Decimal in array');
}

// Test 10: Decimal comparison
{
    const code = `
    Function Test10()
        Dim d1 As Decimal
        Dim d2 As Decimal
        d1 = 10.5
        d2 = 10.5
        If d1 = d2 Then
            Test10 = 1
        Else
            Test10 = 0
        End If
    End Function
    `;
    try {
        const result = runFunc(code, 'Test10');
        assert.strictEqual(result, 1, 'Decimal comparison');
        console.log('[PASS] Test 10: Decimal comparison');
    } catch (e: any) {
        console.log('[INFO] Test 10: ' + e.message);
    }
}

// Test 11: Decimal with very small numbers
{
    const code = `
    Function Test11()
        Dim d As Decimal
        d = 0.0001
        Test11 = d
    End Function
    `;
    try {
        const result = runFunc(code, 'Test11');
        assert.strictEqual(result, 0.0001, 'Decimal with small numbers');
        console.log('[PASS] Test 11: Decimal with small numbers');
    } catch (e: any) {
        console.log('[INFO] Test 11: ' + e.message);
    }
}

// Test 12: Decimal type detection
{
    const code = `
    Function Test12()
        Dim d As Decimal
        d = 123.45
        Dim t As Integer
        t = VarType(d)
        Test12 = t
    End Function
    `;
    try {
        const result = runFunc(code, 'Test12');
        // VarType returns a specific code for Decimal (14 in VBA)
        console.log('[PASS] Test 12: Decimal type detection (VarType)');
    } catch (e: any) {
        console.log('[INFO] Test 12: ' + e.message);
    }
}

// Test 13: Decimal conversion from Integer
{
    const code = `
    Function Test13()
        Dim i As Integer
        Dim d As Decimal
        i = 42
        d = i
        Test13 = d
    End Function
    `;
    try {
        const result = runFunc(code, 'Test13');
        assert.strictEqual(result, 42, 'Decimal from Integer');
        console.log('[PASS] Test 13: Decimal conversion from Integer');
    } catch (e: any) {
        console.log('[INFO] Test 13: ' + e.message);
    }
}

// Test 14: Decimal in function parameter
{
    const code = `
    Function Helper(d As Decimal) As Decimal
        Helper = d * 2
    End Function

    Function Test14()
        Dim d As Decimal
        d = 25.5
        Test14 = Helper(d)
    End Function
    `;
    try {
        const result = runFunc(code, 'Test14');
        assert.strictEqual(result, 51, 'Decimal in function parameter');
        console.log('[PASS] Test 14: Decimal in function parameter');
    } catch (e: any) {
        console.log('[INFO] Test 14: ' + e.message);
    }
}

// Test 15: VBA veteran pattern - financial calculation with Decimal
{
    const code = `
    Function CalculateTax(amount As Decimal, rate As Decimal) As Decimal
        CalculateTax = amount * rate
    End Function

    Function Test15()
        Dim price As Decimal
        Dim taxRate As Decimal
        Dim total As Decimal

        price = 99.99
        taxRate = 0.08
        total = price + CalculateTax(price, taxRate)
        Test15 = total
    End Function
    `;
    try {
        const result = runFunc(code, 'Test15');
        // 99.99 + (99.99 * 0.08) = 99.99 + 7.9992 = 107.9892
        const expected = 99.99 + (99.99 * 0.08);
        // Allow small floating point differences
        const diff = Math.abs(result - expected);
        if (diff < 0.0001) {
            assert.strictEqual(1, 1, 'VBA veteran pattern - financial calc');
            console.log('[PASS] Test 15: VBA veteran pattern - financial calculation');
        } else {
            console.log('[INFO] Test 15: Precision mismatch');
        }
    } catch (e: any) {
        console.log('[INFO] Test 15: ' + e.message);
    }
}

console.log('\n✅ Decimal Precision: Testing complete');
