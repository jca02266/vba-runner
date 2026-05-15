' Currency Type Operations Specification Test
' Tests the 64-bit Currency type with fixed 4-decimal precision

Option Explicit

' Test 1: Basic Currency arithmetic
Sub Test_CurrencyArithmetic(assert)
    Dim c1 As Currency
    Dim c2 As Currency
    Dim result As Currency

    c1 = 100.25
    c2 = 50.75
    result = c1 + c2

    assert.Assert result, 151#, "Currency arithmetic"
End Sub

' Test 2: Currency multiplication with precision
Sub Test_CurrencyMultiplication(assert)
    Dim price As Currency
    Dim quantity As Currency
    Dim total As Currency

    price = 19.99
    quantity = 3
    total = price * quantity

    assert.Assert total, 59.97, "Currency multiplication"
End Sub

' Test 3: Currency division
Sub Test_CurrencyDivision(assert)
    Dim total As Currency
    Dim count As Currency
    Dim average As Currency

    total = 100#
    count = 4
    average = total / count

    assert.Assert average, 25#, "Currency division"
End Sub

' Test 4: Currency with tax calculation
Sub Test_CurrencyTaxCalculation(assert)
    Dim subtotal As Currency
    Dim taxRate As Currency
    Dim tax As Currency
    Dim total As Currency

    subtotal = 100#
    taxRate = 0.08
    tax = subtotal * taxRate
    total = subtotal + tax

    assert.Assert total, 108#, "Currency tax calculation"
End Sub

' Test 5: Currency array operations
Sub Test_CurrencyArray(assert)
    Dim amounts(1 To 3) As Currency
    Dim sum As Currency

    amounts(1) = 10.50
    amounts(2) = 20.75
    amounts(3) = 30.25

    sum = amounts(1) + amounts(2) + amounts(3)
    assert.Assert sum, 61.5, "Currency array operations"
End Sub

' Test 6: Currency type conversion
Sub Test_CurrencyConversion(assert)
    Dim i As Integer
    Dim c As Currency

    i = 42
    c = CCur(i)

    assert.Assert c, 42#, "Currency conversion"
End Sub

' Test 7: Currency comparison
Sub Test_CurrencyComparison(assert)
    Dim c1 As Currency
    Dim c2 As Currency

    c1 = 99.99
    c2 = 99.99

    assert.Assert c1, c2, "Currency comparison"
End Sub

' Test 8: Currency with negative values
Sub Test_CurrencyNegative(assert)
    Dim expense As Currency
    Dim income As Currency
    Dim net As Currency

    expense = -50.25
    income = 100#
    net = income + expense

    assert.Assert net, 49.75, "Currency negative values"
End Sub

' Test 9: Currency precision with 4 decimals
Sub Test_CurrencyPrecision(assert)
    Dim c As Currency
    c = 10.1234
    assert.Assert c, 10.1234, "Currency precision"
End Sub

' Test 10: Currency in function parameter
Function CalculateDiscount(price As Currency, discountPercent As Currency) As Currency
    CalculateDiscount = price * (1 - discountPercent / 100)
End Function

Sub Test_CurrencyFunctionParameter(assert)
    Dim originalPrice As Currency
    Dim discountedPrice As Currency

    originalPrice = 100#
    discountedPrice = CalculateDiscount(originalPrice, 10)

    assert.Assert discountedPrice, 90#, "Currency function parameter"
End Sub
