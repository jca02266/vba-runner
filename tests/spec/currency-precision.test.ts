import { evalVBASingle, assert } from '../../test-libs/test-runner';

function evalVBA(code: string): any {
    return evalVBASingle(code);
}

function runFunc(code: string, name: string, args: any[] = []): any {
    return evalVBA(code).callProcedure(name, args);
}

// Helper: compare Currency result as string (VbaCurrency.toString() or Number.toString())
function cur(v: any): string { return String(v); }

// Test 1: Basic Currency assignment
{
    const code = `
    Function Test1()
        Dim c As Currency
        c = 100.50
        Test1 = c
    End Function
    `;
    const result = runFunc(code, 'Test1');
    assert.strictEqual(cur(result), '100.5', 'Currency basic assignment');
    console.log('[PASS] Test 1: Basic Currency assignment');
}

// Test 2: Currency with CCur() function
{
    const code = `
    Function Test2()
        Dim c As Currency
        c = CCur("250.75")
        Test2 = c
    End Function
    `;
    const result = runFunc(code, 'Test2');
    assert.strictEqual(cur(result), '250.75', 'Currency with CCur()');
    console.log('[PASS] Test 2: CCur() conversion');
}

// Test 3: Currency arithmetic - addition
{
    const code = `
    Function Test3()
        Dim c1 As Currency
        Dim c2 As Currency
        c1 = 150.25
        c2 = 249.75
        Test3 = c1 + c2
    End Function
    `;
    const result = runFunc(code, 'Test3');
    assert.strictEqual(cur(result), '400', 'Currency addition');
    console.log('[PASS] Test 3: Currency addition');
}

// Test 4: Currency arithmetic - subtraction
{
    const code = `
    Function Test4()
        Dim c1 As Currency
        Dim c2 As Currency
        c1 = 500.00
        c2 = 123.45
        Test4 = c1 - c2
    End Function
    `;
    const result = runFunc(code, 'Test4');
    assert.strictEqual(cur(result), '376.55', 'Currency subtraction');
    console.log('[PASS] Test 4: Currency subtraction');
}

// Test 5: Currency arithmetic - multiplication
{
    const code = `
    Function Test5()
        Dim c1 As Currency
        Dim c2 As Currency
        c1 = 50.00
        c2 = 2.5
        Test5 = c1 * c2
    End Function
    `;
    const result = runFunc(code, 'Test5');
    assert.strictEqual(cur(result), '125', 'Currency multiplication');
    console.log('[PASS] Test 5: Currency multiplication');
}

// Test 6: Currency arithmetic - division returns Double
{
    const code = `
    Function Test6()
        Dim c1 As Currency
        Dim c2 As Currency
        c1 = 100.00
        c2 = 4.00
        Test6 = c1 / c2
    End Function
    `;
    const result = runFunc(code, 'Test6');
    assert.strictEqual(result, 25, 'Currency division (Double result)');
    console.log('[PASS] Test 6: Currency division');
}

// Test 7: Currency with 4 decimal places precision
{
    const code = `
    Function Test7()
        Dim c As Currency
        c = 10.1234
        Test7 = c
    End Function
    `;
    const result = runFunc(code, 'Test7');
    assert.strictEqual(cur(result), '10.1234', 'Currency 4 decimal places');
    console.log('[PASS] Test 7: Currency 4 decimal places');
}

// Test 8: Currency with negative numbers
{
    const code = `
    Function Test8()
        Dim c As Currency
        c = -299.99
        Test8 = c
    End Function
    `;
    const result = runFunc(code, 'Test8');
    assert.strictEqual(cur(result), '-299.99', 'Currency negative');
    console.log('[PASS] Test 8: Currency with negative numbers');
}

// Test 9: Currency in array (array elements not typed, comparison as-is)
{
    const code = `
    Function Test9()
        Dim c(1 To 3) As Currency
        c(1) = 10.25
        c(2) = 20.50
        c(3) = 30.75
        Test9 = c(1) + c(2) + c(3)
    End Function
    `;
    const result = runFunc(code, 'Test9');
    assert.strictEqual(cur(result), '61.5', 'Currency in array');
    console.log('[PASS] Test 9: Currency in array');
}

// Test 10: Currency comparison
{
    const code = `
    Function Test10()
        Dim c1 As Currency
        Dim c2 As Currency
        c1 = 99.99
        c2 = 99.99
        If c1 = c2 Then
            Test10 = 1
        Else
            Test10 = 0
        End If
    End Function
    `;
    const result = runFunc(code, 'Test10');
    assert.strictEqual(result, 1, 'Currency comparison');
    console.log('[PASS] Test 10: Currency comparison');
}

// Test 11: Currency type detection
{
    const code = `
    Function Test11()
        Dim c As Currency
        c = 500.00
        Dim t As Integer
        t = VarType(c)
        Test11 = t
    End Function
    `;
    const result = runFunc(code, 'Test11');
    assert.strictEqual(result, 6, 'Currency VarType');
    console.log('[PASS] Test 11: Currency type detection');
}

// Test 12: Currency conversion from Integer
{
    const code = `
    Function Test12()
        Dim i As Integer
        Dim c As Currency
        i = 1000
        c = i
        Test12 = c
    End Function
    `;
    const result = runFunc(code, 'Test12');
    assert.strictEqual(cur(result), '1000', 'Currency from Integer');
    console.log('[PASS] Test 12: Currency from Integer');
}

// Test 13: Currency in function parameter
{
    const code = `
    Function ApplyDiscount(price As Currency, discountRate As Currency) As Currency
        ApplyDiscount = price * (1 - discountRate)
    End Function

    Function Test13()
        Dim originalPrice As Currency
        Dim discount As Currency
        originalPrice = 100.00
        discount = 0.10
        Test13 = ApplyDiscount(originalPrice, discount)
    End Function
    `;
    const result = runFunc(code, 'Test13');
    assert.strictEqual(cur(result), '90', 'Currency in function parameter');
    console.log('[PASS] Test 13: Currency in function parameter');
}

// Test 14: Currency with fractional cents
{
    const code = `
    Function Test14()
        Dim c As Currency
        c = 19.9999
        Test14 = c
    End Function
    `;
    const result = runFunc(code, 'Test14');
    assert.strictEqual(cur(result), '19.9999', 'Currency fractional cents');
    console.log('[PASS] Test 14: Currency with fractional cents');
}

// Test 15: VBA veteran pattern - invoice calculation with Currency
{
    const code = `
    Function CalculateInvoiceTotal(subtotal As Currency, taxRate As Currency, discountPercent As Currency) As Currency
        Dim afterDiscount As Currency
        Dim tax As Currency

        afterDiscount = subtotal * (1 - discountPercent / 100)
        tax = afterDiscount * taxRate
        CalculateInvoiceTotal = afterDiscount + tax
    End Function

    Function Test15()
        Dim total As Currency
        total = CalculateInvoiceTotal(1000.00, 0.08, 10)
        Test15 = total
    End Function
    `;
    const result = runFunc(code, 'Test15');
    // 1000 * (1 - 10/100) = 900; 900 * 0.08 = 72; 900 + 72 = 972
    assert.strictEqual(cur(result), '972', 'Currency invoice calculation');
    console.log('[PASS] Test 15: VBA veteran pattern - invoice calculation');
}

console.log('\n✅ Currency Precision: 全テスト通過');
