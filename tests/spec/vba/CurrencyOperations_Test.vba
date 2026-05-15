' Currency Type Operations Specification Test
' Tests the 64-bit Currency type with fixed 4-decimal precision

Option Explicit

' Test 1: Basic Currency arithmetic
Sub Test_CurrencyArithmetic()
    Dim c1 As Currency
    Dim c2 As Currency
    Dim result As Currency

    c1 = 100.25
    c2 = 50.75
    result = c1 + c2

    Debug.Assert result = 151#, "Currency arithmetic failed"
End Sub

' Test 2: Currency multiplication with precision
Sub Test_CurrencyMultiplication()
    Dim price As Currency
    Dim quantity As Currency
    Dim total As Currency

    price = 19.99
    quantity = 3
    total = price * quantity

    Debug.Assert total = 59.97, "Currency multiplication failed"
End Sub

' Test 3: Currency division
Sub Test_CurrencyDivision()
    Dim total As Currency
    Dim count As Currency
    Dim average As Currency

    total = 100#
    count = 4
    average = total / count

    Debug.Assert average = 25#, "Currency division failed"
End Sub

' Test 4: Currency with tax calculation
Sub Test_CurrencyTaxCalculation()
    Dim subtotal As Currency
    Dim taxRate As Currency
    Dim tax As Currency
    Dim total As Currency

    subtotal = 100#
    taxRate = 0.08
    tax = subtotal * taxRate
    total = subtotal + tax

    Debug.Assert total = 108#, "Currency tax calculation failed"
End Sub

' Test 5: Currency array operations
Sub Test_CurrencyArray()
    Dim amounts(1 To 3) As Currency
    Dim sum As Currency

    amounts(1) = 10.50
    amounts(2) = 20.75
    amounts(3) = 30.25

    sum = amounts(1) + amounts(2) + amounts(3)
    Debug.Assert sum = 61.5, "Currency array operations failed"
End Sub

' Test 6: Currency type conversion
Sub Test_CurrencyConversion()
    Dim i As Integer
    Dim c As Currency

    i = 42
    c = CCur(i)

    Debug.Assert c = 42#, "Currency conversion failed"
End Sub

' Test 7: Currency comparison
Sub Test_CurrencyComparison()
    Dim c1 As Currency
    Dim c2 As Currency

    c1 = 99.99
    c2 = 99.99

    Debug.Assert c1 = c2, "Currency comparison failed"
End Sub

' Test 8: Currency with negative values
Sub Test_CurrencyNegative()
    Dim expense As Currency
    Dim income As Currency
    Dim net As Currency

    expense = -50.25
    income = 100#
    net = income + expense

    Debug.Assert net = 49.75, "Currency negative values failed"
End Sub

' Test 9: Currency precision with 4 decimals
Sub Test_CurrencyPrecision()
    Dim c As Currency
    c = 10.1234
    Debug.Assert c = 10.1234, "Currency precision failed"
End Sub

' Test 10: Currency in function parameter
Function CalculateDiscount(price As Currency, discountPercent As Currency) As Currency
    CalculateDiscount = price * (1 - discountPercent / 100)
End Function

Sub Test_CurrencyFunctionParameter()
    Dim originalPrice As Currency
    Dim discountedPrice As Currency

    originalPrice = 100#
    discountedPrice = CalculateDiscount(originalPrice, 10)

    Debug.Assert discountedPrice = 90#, "Currency function parameter failed"
End Sub
