import { evalVBASingle, assert } from '../../test-libs/test-runner';

function runFunc(code: string, name: string, args: any[] = []): any {
    return evalVBASingle(code).callProcedure(name, args);
}

// Helper: compare VbaDecimal result as string
function dec(v: any): string { return String(v); }

// Test 1: 28-digit precision via CDec(string)
{
    const code = `
    Function Test1()
        Dim d As Decimal
        d = CDec("0.1000000000000000000000000001")
        Test1 = d
    End Function
    `;
    const result = runFunc(code, 'Test1');
    assert.strictEqual(dec(result), '0.1000000000000000000000000001', 'CDec string 28-digit precision');
    console.log('[PASS] Test 1: 28-digit string precision');
}

// Test 2: CDec(0.1) + CDec(0.2) = 0.3 (float-error-free)
{
    const code = `
    Function Test2()
        Test2 = CDec(0.1) + CDec(0.2)
    End Function
    `;
    const result = runFunc(code, 'Test2');
    assert.strictEqual(dec(result), '0.3', 'CDec(0.1) + CDec(0.2) precision');
    console.log('[PASS] Test 2: CDec float-error-free addition');
}

// Test 3: CDec(1) / CDec(3) = 28 threes
{
    const code = `
    Function Test3()
        Test3 = CDec(1) / CDec(3)
    End Function
    `;
    const result = runFunc(code, 'Test3');
    assert.strictEqual(dec(result), '0.3333333333333333333333333333', 'CDec division 28-digit');
    console.log('[PASS] Test 3: CDec(1)/CDec(3) 28-digit division');
}

// Test 4: CDec(10) / CDec(3) = 3.333...
{
    const code = `
    Function Test4()
        Test4 = CDec(10) / CDec(3)
    End Function
    `;
    const result = runFunc(code, 'Test4');
    assert.strictEqual(dec(result), '3.3333333333333333333333333333', 'CDec division integer+frac 28-digit');
    console.log('[PASS] Test 4: CDec(10)/CDec(3) 28-digit division');
}

// Test 5: Exact integer division
{
    const code = `
    Function Test5()
        Test5 = CDec(100) / CDec(4)
    End Function
    `;
    const result = runFunc(code, 'Test5');
    assert.strictEqual(dec(result), '25', 'CDec exact division');
    console.log('[PASS] Test 5: CDec exact division');
}

// Test 6: 28-digit number round-trip
{
    const code = `
    Function Test6()
        Dim d As Decimal
        d = CDec("79228162514264337593543950335")
        Test6 = d
    End Function
    `;
    const result = runFunc(code, 'Test6');
    assert.strictEqual(dec(result), '79228162514264337593543950335', 'CDec MAX_MANTISSA round-trip');
    console.log('[PASS] Test 6: MAX_MANTISSA round-trip');
}

// Test 7: Negative 28-digit precision
{
    const code = `
    Function Test7()
        Dim d As Decimal
        d = CDec("-0.1000000000000000000000000001")
        Test7 = d
    End Function
    `;
    const result = runFunc(code, 'Test7');
    assert.strictEqual(dec(result), '-0.1000000000000000000000000001', 'Negative 28-digit precision');
    console.log('[PASS] Test 7: Negative 28-digit precision');
}

// Test 8: CDec multiplication with scale sum
{
    const code = `
    Function Test8()
        Test8 = CDec("0.0000000001") * CDec("0.0000000001")
    End Function
    `;
    const result = runFunc(code, 'Test8');
    // 10^-10 * 10^-10 = 10^-20 → scale=20, mantissa=1
    assert.strictEqual(dec(result), '0.00000000000000000001', 'CDec multiplication scale sum');
    console.log('[PASS] Test 8: CDec multiplication scale sum');
}

// Test 9: CDec addition with scale alignment
{
    const code = `
    Function Test9()
        Test9 = CDec("0.0000000000000000000000000001") + CDec("1")
    End Function
    `;
    const result = runFunc(code, 'Test9');
    assert.strictEqual(dec(result), '1.0000000000000000000000000001', 'CDec addition scale alignment');
    console.log('[PASS] Test 9: CDec addition scale alignment');
}

// Test 10: CDec subtraction exact
{
    const code = `
    Function Test10()
        Test10 = CDec("1.0000000000000000000000000001") - CDec("1")
    End Function
    `;
    const result = runFunc(code, 'Test10');
    assert.strictEqual(dec(result), '0.0000000000000000000000000001', 'CDec subtraction exact');
    console.log('[PASS] Test 10: CDec subtraction exact');
}

// Test 11: VarType = 14 (Decimal)
{
    const code = `
    Function Test11()
        Test11 = VarType(CDec(1))
    End Function
    `;
    const result = runFunc(code, 'Test11');
    assert.strictEqual(result, 14, 'VarType(CDec) = 14');
    console.log('[PASS] Test 11: VarType(CDec) = 14');
}

// Test 12: TypeName = "Decimal"
{
    const code = `
    Function Test12()
        Test12 = TypeName(CDec(1))
    End Function
    `;
    const result = runFunc(code, 'Test12');
    assert.strictEqual(result, 'Decimal', 'TypeName(CDec) = "Decimal"');
    console.log('[PASS] Test 12: TypeName(CDec) = "Decimal"');
}

// Test 13: Overflow beyond MAX_MANTISSA
{
    const code = `
    Function Test13()
        On Error GoTo errHandler
        Dim d As Decimal
        d = CDec("79228162514264337593543950336")
        Test13 = 0
        Exit Function
    errHandler:
        Test13 = Err.Number
    End Function
    `;
    const result = runFunc(code, 'Test13');
    assert.strictEqual(result, 6, 'CDec overflow Error 6');
    console.log('[PASS] Test 13: Overflow Error 6');
}

// Test 14: CDec from Currency preserves 4 decimal places
{
    const code = `
    Function Test14()
        Dim c As Currency
        c = CCur("1.2345")
        Test14 = CDec(c)
    End Function
    `;
    const result = runFunc(code, 'Test14');
    assert.strictEqual(dec(result), '1.2345', 'CDec from Currency');
    console.log('[PASS] Test 14: CDec from Currency');
}

// Test 15: Array of Decimal with 28-digit CDec value
{
    const code = `
    Function Test15()
        Dim d(1 To 2) As Decimal
        d(1) = CDec("0.1000000000000000000000000001")
        d(2) = CDec("0.2000000000000000000000000002")
        Test15 = d(1) + d(2)
    End Function
    `;
    const result = runFunc(code, 'Test15');
    assert.strictEqual(dec(result), '0.3000000000000000000000000003', 'Array Decimal 28-digit addition');
    console.log('[PASS] Test 15: Array Decimal 28-digit addition');
}

// Test 16: Unary minus on CDec
{
    const code = `
    Function Test16()
        Dim d As Decimal
        d = CDec("0.1000000000000000000000000001")
        Test16 = -d
    End Function
    `;
    const result = runFunc(code, 'Test16');
    assert.strictEqual(dec(result), '-0.1000000000000000000000000001', 'Unary minus on CDec');
    console.log('[PASS] Test 16: Unary minus on CDec');
}

// Test 17: CDec(1) / CDec(7) banker's rounding at 29th digit
{
    const code = `
    Function Test17()
        Test17 = CDec(1) / CDec(7)
    End Function
    `;
    const result = runFunc(code, 'Test17');
    // 1/7 = 0.142857142857... → 28 digits with banker's rounding
    // 10^28 / 7 = 1428571428571428571428571428, remainder 4 → 4/7 > 0.5 → round up
    assert.strictEqual(dec(result), '0.1428571428571428571428571429', 'CDec(1)/CDec(7) 28-digit');
    console.log('[PASS] Test 17: CDec(1)/CDec(7) 28-digit');
}

// Test 18: CStr of CDec preserves trailing zeros only when significant
{
    const code = `
    Function Test18()
        Dim d As Decimal
        d = CDec("3.14")
        Test18 = CStr(d)
    End Function
    `;
    const result = runFunc(code, 'Test18');
    assert.strictEqual(result, '3.14', 'CStr(CDec) no spurious trailing zeros');
    console.log('[PASS] Test 18: CStr of CDec no spurious trailing zeros');
}

console.log('\n✅ Decimal BigInt (Phase 3): 全テスト通過');
