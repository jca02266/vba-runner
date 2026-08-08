/**
 * Select Case Statement のテスト (§5.4.2.10)
 *
 * 仕様書の range-clause 3種を網羅:
 *   1. expression        : Case 1, 2, 3
 *   2. start To end      : Case 1 To 5
 *   3. [Is] comparison   : Case Is > 10
 */
import { evalVBASingle, assert, vbaTrue, vbaFalse } from '../../test-libs/test-runner';

function evalVBA(code: string): any {
    return evalVBASingle(code);
}

function runFunc(code: string, name: string, args: any[] = []): any {
    return evalVBA(code).callProcedure(name, args);
}

// --- 1. 単純な値マッチ ---
const basicCode = `
Function GetLabel(x)
    Select Case x
        Case 1
            GetLabel = "one"
        Case 2
            GetLabel = "two"
        Case 3
            GetLabel = "three"
        Case Else
            GetLabel = "other"
    End Select
End Function
`;
assert.strictEqual(runFunc(basicCode, 'GetLabel', [1]), 'one',   'Case 1');
assert.strictEqual(runFunc(basicCode, 'GetLabel', [2]), 'two',   'Case 2');
assert.strictEqual(runFunc(basicCode, 'GetLabel', [3]), 'three', 'Case 3');
assert.strictEqual(runFunc(basicCode, 'GetLabel', [9]), 'other', 'Case Else');
console.log('[PASS] 単純な値マッチ');

// --- 2. カンマ区切り複数値 ---
const multiCode = `
Function IsWeekend(day)
    Select Case day
        Case 1, 7
            IsWeekend = True
        Case Else
            IsWeekend = False
    End Select
End Function
`;
assert.strictEqual(runFunc(multiCode, 'IsWeekend', [1]), vbaTrue, 'day=1 is weekend');
assert.strictEqual(runFunc(multiCode, 'IsWeekend', [7]), vbaTrue, 'day=7 is weekend');
assert.strictEqual(runFunc(multiCode, 'IsWeekend', [3]), vbaFalse, 'day=3 is not weekend');
console.log('[PASS] カンマ区切り複数値');

// --- 3. To による範囲マッチ ---
const rangeCode = `
Function GetGrade(score)
    Select Case score
        Case 90 To 100
            GetGrade = "A"
        Case 70 To 89
            GetGrade = "B"
        Case 50 To 69
            GetGrade = "C"
        Case Else
            GetGrade = "F"
    End Select
End Function
`;
assert.strictEqual(runFunc(rangeCode, 'GetGrade', [95]),  'A', 'score 95 -> A');
assert.strictEqual(runFunc(rangeCode, 'GetGrade', [90]),  'A', 'score 90 -> A (境界)');
assert.strictEqual(runFunc(rangeCode, 'GetGrade', [100]), 'A', 'score 100 -> A (境界)');
assert.strictEqual(runFunc(rangeCode, 'GetGrade', [75]),  'B', 'score 75 -> B');
assert.strictEqual(runFunc(rangeCode, 'GetGrade', [55]),  'C', 'score 55 -> C');
assert.strictEqual(runFunc(rangeCode, 'GetGrade', [40]),  'F', 'score 40 -> F');
console.log('[PASS] To による範囲マッチ');

// --- 4. Is による比較マッチ ---
const isCode = `
Function Classify(n)
    Select Case n
        Case Is < 0
            Classify = "negative"
        Case Is = 0
            Classify = "zero"
        Case Is > 0
            Classify = "positive"
    End Select
End Function
`;
assert.strictEqual(runFunc(isCode, 'Classify', [-5]), 'negative', 'Is < 0');
assert.strictEqual(runFunc(isCode, 'Classify', [0]),  'zero',     'Is = 0');
assert.strictEqual(runFunc(isCode, 'Classify', [3]),  'positive', 'Is > 0');
console.log('[PASS] Is による比較マッチ');

// --- 5. 最初にマッチした Case のみ実行される（フォールスルーなし）---
const firstMatchCode = `
Function FirstMatch(x)
    Select Case x
        Case 1 To 10
            FirstMatch = "low"
        Case 5 To 20
            FirstMatch = "high"
    End Select
End Function
`;
assert.strictEqual(runFunc(firstMatchCode, 'FirstMatch', [5]), 'low', '最初のマッチのみ実行');
console.log('[PASS] 最初にマッチした Case のみ実行');

// --- 6. Case Else がない場合にマッチしなければ何もしない ---
const noElseCode = `
Function NoElse(x)
    NoElse = "default"
    Select Case x
        Case 1
            NoElse = "one"
    End Select
End Function
`;
assert.strictEqual(runFunc(noElseCode, 'NoElse', [1]), 'one',     'マッチあり');
assert.strictEqual(runFunc(noElseCode, 'NoElse', [2]), 'default', 'マッチなし、デフォルト値を維持');
console.log('[PASS] Case Else なし');

// --- 7. 文字列マッチ ---
const strCode = `
Function DayType(day)
    Select Case day
        Case "Mon", "Tue", "Wed", "Thu", "Fri"
            DayType = "weekday"
        Case "Sat", "Sun"
            DayType = "weekend"
        Case Else
            DayType = "unknown"
    End Select
End Function
`;
assert.strictEqual(runFunc(strCode, 'DayType', ['Mon']), 'weekday', 'Mon is weekday');
assert.strictEqual(runFunc(strCode, 'DayType', ['Sat']), 'weekend', 'Sat is weekend');
assert.strictEqual(runFunc(strCode, 'DayType', ['xyz']), 'unknown', 'unknown day');
console.log('[PASS] 文字列マッチ');

console.log('\n✅ Select Case: 全テスト通過');

// Bug 191-A: Empty follows the comparison context in Select Case equality.
{
    const result = runFunc(`
Function TestEmptyCase() As String
    Dim value As Variant
    value = Empty
    Select Case value
        Case 0
            TestEmptyCase = "zero"
        Case Else
            TestEmptyCase = "else"
    End Select
End Function
`, 'TestEmptyCase');
    assert.strictEqual(result, 'zero', 'Empty matches numeric Case 0');

    const stringResult = runFunc(`
Function TestEmptyStringCase() As String
    Dim value As Variant
    value = Empty
    Select Case value
        Case ""
            TestEmptyStringCase = "empty"
        Case Else
            TestEmptyStringCase = "else"
    End Select
End Function
`, 'TestEmptyStringCase');
    assert.strictEqual(stringResult, 'empty', 'Empty matches string Case ""');

    const notEqualResult = runFunc(`
Function TestEmptyNotEqual() As String
    Dim value As Variant
    value = Empty
    Select Case value
        Case Is <> 0
            TestEmptyNotEqual = "wrong"
        Case Else
            TestEmptyNotEqual = "zero"
    End Select
End Function
`, 'TestEmptyNotEqual');
    assert.strictEqual(notEqualResult, 'zero', 'Empty Case Is <> 0 is false');
    console.log('[PASS] Bug 191-A: Select Case Empty coercion');
}

// Bug 192-A: Decimal equality must not round through a JS Number.
{
    const result = runFunc(`
Function TestDecimalPrecisionCase() As String
    Dim value As Variant
    value = CDec("12345678901234567.1")
    Select Case value
        Case CDec("12345678901234567.2")
            TestDecimalPrecisionCase = "wrong-equal"
        Case Else
            TestDecimalPrecisionCase = "distinct"
    End Select
End Function
`, 'TestDecimalPrecisionCase');
    assert.strictEqual(result, 'distinct', 'Select Case preserves Decimal equality precision');
    console.log('[PASS] Bug 192-A: Select Case Decimal precision');

    const rangeResult = runFunc(`
Function TestDecimalRange() As String
    Dim value As Variant
    value = CDec("9.5")
    Select Case value
        Case CDec("9") To CDec("10")
            TestDecimalRange = "range"
        Case Else
            TestDecimalRange = "no"
    End Select
End Function
`, 'TestDecimalRange');
    assert.strictEqual(rangeResult, 'range', 'Select Case preserves Decimal range comparison');

    const currencyRange = runFunc(`
Function TestCurrencyRange() As String
    Dim value As Currency
    value = CCur("9.5")
    Select Case value
        Case CCur(9) To CCur(10)
            TestCurrencyRange = "range"
        Case Else
            TestCurrencyRange = "no"
    End Select
End Function
`, 'TestCurrencyRange');
    assert.strictEqual(currencyRange, 'range', 'Select Case preserves Currency range comparison');

    const relationalPrecision = runFunc(`
Function TestDecimalRelationalPrecision() As String
    Dim value As Variant
    value = CDec("123456789012345.1")
    Select Case value
        Case Is < CDec("123456789012345.2")
            TestDecimalRelationalPrecision = "less"
        Case Else
            TestDecimalRelationalPrecision = "wrong"
    End Select
End Function
`, 'TestDecimalRelationalPrecision');
    assert.strictEqual(relationalPrecision, 'less', 'Select Case relational Decimal comparison preserves precision');
    console.log('[PASS] Bug 193-A: Select Case Decimal/Currency ranges');

    const decimalString = runFunc(`
Function TestDecimalStringCompare() As String
    Dim value As Variant
    value = CDec("12345678901234567.1")
    Select Case value
        Case Is = "12345678901234567.1"
            TestDecimalStringCompare = "equal"
        Case Else
            TestDecimalStringCompare = "wrong"
    End Select
End Function
`, 'TestDecimalStringCompare');
    assert.strictEqual(decimalString, 'equal', 'Decimal compares exactly with numeric String');

    const currencyString = runFunc(`
Function TestCurrencyStringCompare() As String
    Dim value As Currency
    value = CCur("900719925474.0993")
    Select Case value
        Case Is > "900719925474.0992"
            TestCurrencyStringCompare = "greater"
        Case Else
            TestCurrencyStringCompare = "wrong"
    End Select
End Function
`, 'TestCurrencyStringCompare');
    assert.strictEqual(currencyString, 'greater', 'Currency compares exactly with numeric String');
    console.log('[PASS] Bug 196-A: Select Case fixed-point String comparison');
}

// Bug 194-A: numeric Variant strings use numeric equality in Select Case.
{
    const result = runFunc(`
Function TestNumericStringCase() As String
    Dim value As Variant
    value = "100"
    Select Case value
        Case 100
            TestNumericStringCase = "number"
        Case Else
            TestNumericStringCase = "else"
    End Select
End Function
`, 'TestNumericStringCase');
    assert.strictEqual(result, 'number', 'Numeric Variant String matches numeric Case');
    console.log('[PASS] Bug 194-A: Select Case numeric String equality');
}

// Bug 195-A: Null comparisons do not match or throw a JS Symbol error.
{
    const result = runFunc(`
Function TestNullCase() As String
    Dim value As Variant
    value = Null
    Select Case value
        Case Null
            TestNullCase = "wrong-null"
        Case Is <> 0
            TestNullCase = "wrong-neq"
        Case Is > 0
            TestNullCase = "wrong-gt"
        Case Else
            TestNullCase = "else"
    End Select
End Function
`, 'TestNullCase');
    assert.strictEqual(result, 'else', 'Null Select Case comparisons fall through without throwing');
    console.log('[PASS] Bug 195-A: Select Case Null comparison');
}

// Bug 197-A: a nonnumeric Variant String must propagate Type mismatch when
// compared with a numeric Select Case value, including inverted Is clauses.
{
    const result = runFunc(`
Function TestInvalidStringCase(value As Variant, mode As Integer) As String
    On Error GoTo EH
    If mode = 0 Then
        Select Case value
            Case 1
                TestInvalidStringCase = "wrong"
            Case Else
                TestInvalidStringCase = "else"
        End Select
    Else
        Select Case 1
            Case Is <> value
                TestInvalidStringCase = "wrong"
            Case Else
                TestInvalidStringCase = "else"
        End Select
    End If
    Exit Function
EH:
    TestInvalidStringCase = "err:" & CStr(Err.Number)
End Function
`, 'TestInvalidStringCase', ['abc', 0]);
    const emptyResult = runFunc(`
Function TestInvalidEmptyCase(value As Variant) As String
    On Error GoTo EH
    Select Case value
        Case Is <> 1
            TestInvalidEmptyCase = "wrong"
        Case Else
            TestInvalidEmptyCase = "else"
    End Select
    Exit Function
EH:
    TestInvalidEmptyCase = "err:" & CStr(Err.Number)
End Function
`, 'TestInvalidEmptyCase', ['']);
    const invertedResult = runFunc(`
Function TestInvalidInvertedCase() As String
    On Error GoTo EH
    Select Case 1
        Case Is <> "abc"
            TestInvalidInvertedCase = "wrong"
        Case Else
            TestInvalidInvertedCase = "else"
    End Select
    Exit Function
EH:
    TestInvalidInvertedCase = "err:" & CStr(Err.Number)
End Function
`, 'TestInvalidInvertedCase');
    assert.strictEqual(result, 'err:13', 'Nonnumeric Variant String propagates Type mismatch');
    assert.strictEqual(emptyResult, 'err:13', 'Empty String propagates Type mismatch');
    assert.strictEqual(invertedResult, 'err:13', 'Case Is <> propagates Type mismatch');
    console.log('[PASS] Bug 197-A: Select Case invalid String errors');
}

// Bug 173-A: Select Case の Date 式比較は値で一致判定する
{
    const result = runFunc(`
Function TestDateCase() As String
    Dim d As Date
    d = DateSerial(2024, 1, 2)
    Select Case d
        Case DateSerial(2024, 1, 2)
            TestDateCase = "D"
        Case Else
            TestDateCase = "X"
    End Select
End Function
`, 'TestDateCase');
    assert.strictEqual(result, 'D', 'Select Case はDateオブジェクトをシリアル値で比較する');
    console.log('[PASS] Bug 173-A: Select Case Date equality');
}

// Bug 184-A: LongLong values must match ordinary integral Case expressions.
{
    const result = runFunc(`
Function TestLongLongCase() As String
    Dim value As LongLong
    value = CLngLng(42)
    Select Case value
        Case 42
            TestLongLongCase = "number"
        Case Else
            TestLongLongCase = "else"
    End Select
End Function
`, 'TestLongLongCase');
    assert.strictEqual(result, 'number', 'Select Case compares LongLong with integral numbers');
    console.log('[PASS] Bug 184-A: Select Case LongLong equality');
}

// Bug 198-A: LongLong comparisons preserve digits beyond JavaScript's safe
// integer range when the other operand is a numeric String.
{
    const result = runFunc(`
Function TestLongLongStringBoundary() As String
    Dim value As LongLong
    Dim text As Variant
    value = CLngLng("9007199254740993")
    text = "9007199254740993"
    Select Case value
        Case text
            TestLongLongStringBoundary = "equal"
        Case Is > text
            TestLongLongStringBoundary = "greater"
        Case Is < text
            TestLongLongStringBoundary = "less"
        Case Else
            TestLongLongStringBoundary = "else"
    End Select
End Function
`, 'TestLongLongStringBoundary');
    const range = runFunc(`
Function TestLongLongStringRange() As String
    Dim value As LongLong
    value = CLngLng("9007199254740993")
    Select Case value
        Case "9007199254740992" To "9007199254740994"
            TestLongLongStringRange = "in"
        Case Else
            TestLongLongStringRange = "out"
    End Select
End Function
`, 'TestLongLongStringRange');
    const ordinary = runFunc(`
Function TestLongLongStringEquality() As String
    Dim value As LongLong
    Dim text As String
    value = CLngLng("9007199254740993")
    text = "9007199254740993"
    If value = text Then
        TestLongLongStringEquality = "equal"
    Else
        TestLongLongStringEquality = "else"
    End If
End Function
`, 'TestLongLongStringEquality');
    assert.strictEqual(result, 'equal', 'LongLong Select Case equality preserves 64-bit String digits');
    assert.strictEqual(range, 'in', 'LongLong Select Case range preserves 64-bit String digits');
    assert.strictEqual(ordinary, 'equal', 'LongLong ordinary equality preserves 64-bit String digits');
    console.log('[PASS] Bug 198-A: LongLong numeric String precision');
}
