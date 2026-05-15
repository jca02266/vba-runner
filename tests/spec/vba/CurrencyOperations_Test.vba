' Currency Type Operations Specification Test
' Tests the 64-bit Currency type with fixed 4-decimal precision

Option Explicit

' Test 1: Basic Currency arithmetic
Function TestCurrencyArithmetic() As Boolean
    Dim c1 As Currency
    Dim c2 As Currency
    Dim result As Currency

    c1 = 100.25
    c2 = 50.75
    result = c1 + c2

    TestCurrencyArithmetic = (result = 151#)
End Function

' Test 2: Currency multiplication with precision
Function TestCurrencyMultiplication() As Boolean
    Dim price As Currency
    Dim quantity As Currency
    Dim total As Currency

    price = 19.99
    quantity = 3
    total = price * quantity

    TestCurrencyMultiplication = (total = 59.97)
End Function

' Test 3: Currency division
Function TestCurrencyDivision() As Boolean
    Dim total As Currency
    Dim count As Currency
    Dim average As Currency

    total = 100#
    count = 4
    average = total / count

    TestCurrencyDivision = (average = 25#)
End Function

' Test 4: Currency with tax calculation
Function TestCurrencyTaxCalculation() As Boolean
    Dim subtotal As Currency
    Dim taxRate As Currency
    Dim tax As Currency
    Dim total As Currency

    subtotal = 100#
    taxRate = 0.08
    tax = subtotal * taxRate
    total = subtotal + tax

    TestCurrencyTaxCalculation = (total = 108#)
End Function

' Test 5: Currency array operations
Function TestCurrencyArray() As Boolean
    Dim amounts(1 To 3) As Currency
    Dim sum As Currency

    amounts(1) = 10.50
    amounts(2) = 20.75
    amounts(3) = 30.25

    sum = amounts(1) + amounts(2) + amounts(3)
    TestCurrencyArray = (sum = 61.5)
End Function

' Test 6: Currency type conversion
Function TestCurrencyConversion() As Boolean
    Dim i As Integer
    Dim c As Currency

    i = 42
    c = CCur(i)

    TestCurrencyConversion = (c = 42#)
End Function

' Test 7: Currency comparison
Function TestCurrencyComparison() As Boolean
    Dim c1 As Currency
    Dim c2 As Currency

    c1 = 99.99
    c2 = 99.99

    TestCurrencyComparison = (c1 = c2)
End Function

' Test 8: Currency with negative values
Function TestCurrencyNegative() As Boolean
    Dim expense As Currency
    Dim income As Currency
    Dim net As Currency

    expense = -50.25
    income = 100#
    net = income + expense

    TestCurrencyNegative = (net = 49.75)
End Function

' Test 9: Currency precision with 4 decimals
Function TestCurrencyPrecision() As Boolean
    Dim c As Currency
    c = 10.1234
    TestCurrencyPrecision = (c = 10.1234)
End Function

' Test 10: Currency in function parameter
Function CalculateDiscount(price As Currency, discountPercent As Currency) As Currency
    CalculateDiscount = price * (1 - discountPercent / 100)
End Function

Function TestCurrencyFunctionParameter() As Boolean
    Dim originalPrice As Currency
    Dim discountedPrice As Currency

    originalPrice = 100#
    discountedPrice = CalculateDiscount(originalPrice, 10)

    TestCurrencyFunctionParameter = (discountedPrice = 90#)
End Function

' Main test runner
Function RunAllTests() As Boolean
    Dim allPass As Boolean
    Dim testResults As String

    allPass = True
    testResults = "=== Currency Operations Specification Test ===" & vbCrLf

    ' Run all tests
    If Not TestCurrencyArithmetic Then
        allPass = False
        testResults = testResults & "[FAIL] TestCurrencyArithmetic" & vbCrLf
    Else
        testResults = testResults & "[PASS] TestCurrencyArithmetic" & vbCrLf
    End If

    If Not TestCurrencyMultiplication Then
        allPass = False
        testResults = testResults & "[FAIL] TestCurrencyMultiplication" & vbCrLf
    Else
        testResults = testResults & "[PASS] TestCurrencyMultiplication" & vbCrLf
    End If

    If Not TestCurrencyDivision Then
        allPass = False
        testResults = testResults & "[FAIL] TestCurrencyDivision" & vbCrLf
    Else
        testResults = testResults & "[PASS] TestCurrencyDivision" & vbCrLf
    End If

    If Not TestCurrencyTaxCalculation Then
        allPass = False
        testResults = testResults & "[FAIL] TestCurrencyTaxCalculation" & vbCrLf
    Else
        testResults = testResults & "[PASS] TestCurrencyTaxCalculation" & vbCrLf
    End If

    If Not TestCurrencyArray Then
        allPass = False
        testResults = testResults & "[FAIL] TestCurrencyArray" & vbCrLf
    Else
        testResults = testResults & "[PASS] TestCurrencyArray" & vbCrLf
    End If

    If Not TestCurrencyConversion Then
        allPass = False
        testResults = testResults & "[FAIL] TestCurrencyConversion" & vbCrLf
    Else
        testResults = testResults & "[PASS] TestCurrencyConversion" & vbCrLf
    End If

    If Not TestCurrencyComparison Then
        allPass = False
        testResults = testResults & "[FAIL] TestCurrencyComparison" & vbCrLf
    Else
        testResults = testResults & "[PASS] TestCurrencyComparison" & vbCrLf
    End If

    If Not TestCurrencyNegative Then
        allPass = False
        testResults = testResults & "[FAIL] TestCurrencyNegative" & vbCrLf
    Else
        testResults = testResults & "[PASS] TestCurrencyNegative" & vbCrLf
    End If

    If Not TestCurrencyPrecision Then
        allPass = False
        testResults = testResults & "[FAIL] TestCurrencyPrecision" & vbCrLf
    Else
        testResults = testResults & "[PASS] TestCurrencyPrecision" & vbCrLf
    End If

    If Not TestCurrencyFunctionParameter Then
        allPass = False
        testResults = testResults & "[FAIL] TestCurrencyFunctionParameter" & vbCrLf
    Else
        testResults = testResults & "[PASS] TestCurrencyFunctionParameter" & vbCrLf
    End If

    testResults = testResults & vbCrLf & "=== Test Complete ===" & vbCrLf

    Debug.Print testResults
    RunAllTests = allPass
End Function
