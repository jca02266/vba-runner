import { Evaluator } from '../../src/engine/evaluator';
import { evalVBASingle, assert } from '../../test-libs/test-runner';

function evalVBA(code: string): Evaluator {
    return evalVBASingle(code);
}

function runFunc(code: string, name: string, args: any[] = []): any {
    return evalVBA(code).callProcedure(name, args);
}

// Test 1: Basic 2D array iteration order (row-major, leftmost dimension varies fastest)
{
    const code = `
        Function Test1() As String
            Dim arr(1 To 2, 1 To 3) As Integer
            arr(1, 1) = 11: arr(1, 2) = 12: arr(1, 3) = 13
            arr(2, 1) = 21: arr(2, 2) = 22: arr(2, 3) = 23

            Dim result As String
            result = ""
            Dim elem As Variant
            For Each elem In arr
                result = result & CStr(elem) & ","
            Next
            Test1 = result
        End Function
    `;

    try {
        const result = runFunc(code, 'Test1');
        const expected = '11,21,12,22,13,23,';
        assert.strictEqual(result, expected, '2D array iterates in row-major order (leftmost dimension first)');
        console.log('[PASS] Basic 2D array iteration order');
    } catch (e: any) {
        console.log('[FAIL] Basic 2D array iteration order:', e.message);
    }
}

// Test 2: 2D array with 0-based indexing
{
    const code = `
        Function Test2() As String
            Dim arr(0 To 1, 0 To 2) As Integer
            arr(0, 0) = 10: arr(0, 1) = 11: arr(0, 2) = 12
            arr(1, 0) = 20: arr(1, 1) = 21: arr(1, 2) = 22

            Dim result As String
            result = ""
            Dim elem As Variant
            For Each elem In arr
                result = result & CStr(elem) & ","
            Next
            Test2 = result
        End Function
    `;

    try {
        const result = runFunc(code, 'Test2');
        const expected = '10,20,11,21,12,22,';
        assert.strictEqual(result, expected, '0-based 2D array iteration follows row-major order');
        console.log('[PASS] 2D array with 0-based indexing');
    } catch (e: any) {
        console.log('[FAIL] 2D array with 0-based indexing:', e.message);
    }
}

// Test 3: 2D array with asymmetric bounds
{
    const code = `
        Function Test3() As String
            Dim arr(5 To 6, 10 To 12) As Integer
            arr(5, 10) = 1: arr(5, 11) = 2: arr(5, 12) = 3
            arr(6, 10) = 4: arr(6, 11) = 5: arr(6, 12) = 6

            Dim result As String
            result = ""
            Dim elem As Variant
            For Each elem In arr
                result = result & CStr(elem) & ","
            Next
            Test3 = result
        End Function
    `;

    try {
        const result = runFunc(code, 'Test3');
        const expected = '1,4,2,5,3,6,';
        assert.strictEqual(result, expected, 'Asymmetric bounds preserve row-major order');
        console.log('[PASS] 2D array with asymmetric bounds');
    } catch (e: any) {
        console.log('[FAIL] 2D array with asymmetric bounds:', e.message);
    }
}

// Test 4: 3D array iteration order (leftmost varies fastest)
{
    const code = `
        Function Test4() As String
            Dim arr(1 To 2, 1 To 2, 1 To 2) As Integer
            Dim idx As Integer
            idx = 1
            Dim i As Integer, j As Integer, k As Integer
            For i = 1 To 2
                For j = 1 To 2
                    For k = 1 To 2
                        arr(i, j, k) = idx
                        idx = idx + 1
                    Next
                Next
            Next

            Dim result As String
            result = ""
            Dim elem As Variant
            For Each elem In arr
                result = result & CStr(elem) & ","
            Next
            Test4 = result
        End Function
    `;

    try {
        const result = runFunc(code, 'Test4');
        const expected = '1,5,3,7,2,6,4,8,';
        assert.strictEqual(result, expected, '3D array iterates leftmost dimension first');
        console.log('[PASS] 3D array iteration order');
    } catch (e: any) {
        console.log('[FAIL] 3D array iteration order:', e.message);
    }
}

// Test 5: For Each with array element modification
{
    const code = `
        Function Test5() As Integer
            Dim arr(1 To 2, 1 To 2) As Integer
            arr(1, 1) = 10: arr(1, 2) = 20
            arr(2, 1) = 30: arr(2, 2) = 40

            Dim elem As Variant
            Dim sum As Integer
            sum = 0
            For Each elem In arr
                sum = sum + elem
            Next
            Test5 = sum
        End Function
    `;

    try {
        const result = runFunc(code, 'Test5');
        assert.strictEqual(result, 100, 'For Each can access all 2D array elements for aggregation');
        console.log('[PASS] For Each with array element aggregation');
    } catch (e: any) {
        console.log('[FAIL] For Each with array element aggregation:', e.message);
    }
}

// Test 6: Nested For Each on 2D arrays
{
    const code = `
        Function Test6() As Integer
            Dim arr1(1 To 2, 1 To 2) As Integer
            arr1(1, 1) = 1: arr1(1, 2) = 2
            arr1(2, 1) = 3: arr1(2, 2) = 4

            Dim arr2(1 To 2, 1 To 2) As Integer
            arr2(1, 1) = 5: arr2(1, 2) = 6
            arr2(2, 1) = 7: arr2(2, 2) = 8

            Dim sum As Integer
            sum = 0
            Dim e1 As Variant, e2 As Variant
            For Each e1 In arr1
                For Each e2 In arr2
                    sum = sum + e1 + e2
                Next
            Next
            Test6 = sum
        End Function
    `;

    try {
        const result = runFunc(code, 'Test6');
        const arr1Sum = 1 + 3 + 2 + 4;
        const arr2Sum = 5 + 7 + 6 + 8;
        const expected = arr1Sum * 4 + arr2Sum * 4;
        assert.strictEqual(result, expected, 'Nested For Each loops on 2D arrays work correctly');
        console.log('[PASS] Nested For Each on 2D arrays');
    } catch (e: any) {
        console.log('[FAIL] Nested For Each on 2D arrays:', e.message);
    }
}

// Test 7: 2D array with single row (edge case)
{
    const code = `
        Function Test7() As String
            Dim arr(1 To 1, 1 To 4) As Integer
            arr(1, 1) = 10: arr(1, 2) = 20: arr(1, 3) = 30: arr(1, 4) = 40

            Dim result As String
            result = ""
            Dim elem As Variant
            For Each elem In arr
                result = result & CStr(elem) & ","
            Next
            Test7 = result
        End Function
    `;

    try {
        const result = runFunc(code, 'Test7');
        const expected = '10,20,30,40,';
        assert.strictEqual(result, expected, 'Single-row 2D array iteration works');
        console.log('[PASS] 2D array with single row');
    } catch (e: any) {
        console.log('[FAIL] 2D array with single row:', e.message);
    }
}

// Test 8: 2D array with single column (edge case)
{
    const code = `
        Function Test8() As String
            Dim arr(1 To 4, 1 To 1) As Integer
            arr(1, 1) = 10: arr(2, 1) = 20: arr(3, 1) = 30: arr(4, 1) = 40

            Dim result As String
            result = ""
            Dim elem As Variant
            For Each elem In arr
                result = result & CStr(elem) & ","
            Next
            Test8 = result
        End Function
    `;

    try {
        const result = runFunc(code, 'Test8');
        const expected = '10,20,30,40,';
        assert.strictEqual(result, expected, 'Single-column 2D array iteration works');
        console.log('[PASS] 2D array with single column');
    } catch (e: any) {
        console.log('[FAIL] 2D array with single column:', e.message);
    }
}

// Test 9: Real-world pattern - 2D table processing (row-major access)
{
    const code = `
        Function Test9() As Integer
            Dim table(1 To 3, 1 To 3) As Integer
            table(1, 1) = 1: table(1, 2) = 2: table(1, 3) = 3
            table(2, 1) = 4: table(2, 2) = 5: table(2, 3) = 6
            table(3, 1) = 7: table(3, 2) = 8: table(3, 3) = 9

            Dim maxVal As Integer
            maxVal = 0
            Dim elem As Variant
            For Each elem In table
                If elem > maxVal Then
                    maxVal = elem
                End If
            Next
            Test9 = maxVal
        End Function
    `;

    try {
        const result = runFunc(code, 'Test9');
        assert.strictEqual(result, 9, 'Real-world: Find max value in 2D table');
        console.log('[PASS] Real-world: 2D table processing');
    } catch (e: any) {
        console.log('[FAIL] Real-world: 2D table processing:', e.message);
    }
}

// Test 10: 2D array with mixed string and numeric elements
{
    const code = `
        Function Test10() As String
            Dim arr(1 To 2, 1 To 2) As Variant
            arr(1, 1) = "A": arr(1, 2) = "B"
            arr(2, 1) = "C": arr(2, 2) = "D"

            Dim result As String
            result = ""
            Dim elem As Variant
            For Each elem In arr
                result = result & elem
            Next
            Test10 = result
        End Function
    `;

    try {
        const result = runFunc(code, 'Test10');
        const expected = 'ACBD';
        assert.strictEqual(result, expected, 'Variant 2D array preserves element types during iteration');
        console.log('[PASS] 2D Variant array with mixed types');
    } catch (e: any) {
        console.log('[FAIL] 2D Variant array with mixed types:', e.message);
    }
}

console.log('\n✅ 2D Array For Each: 全テスト完了');
