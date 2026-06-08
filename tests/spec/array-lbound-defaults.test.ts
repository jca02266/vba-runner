import { Evaluator } from '../../src/engine/evaluator';
import { evalVBASingle, assert } from '../../test-libs/test-runner';

function evalVBA(code: string): Evaluator {
    return evalVBASingle(code);
}

function runFunc(code: string, name: string, args: any[] = []): any {
    return evalVBA(code).callProcedure(name, args);
}

// Test 1: Array() function creates 0-based array
{
    const code = String.raw`
        Function TestArrayLBound()
            Dim arr
            arr = Array(10, 20, 30)
            TestArrayLBound = LBound(arr)
        End Function
    `;

    try {
        const result = runFunc(code, 'TestArrayLBound');
        assert.strictEqual(result, 0, 'Array() creates 0-based array (LBound = 0)');
        console.log('[PASS] Array() LBound = 0');
    } catch (e: any) {
        console.log('[FAIL] Array() LBound = 0:', e.message);
    }
}

// Test 2: Array() UBound is elements - 1
{
    const code = String.raw`
        Function TestArrayUBound()
            Dim arr
            arr = Array("a", "b", "c", "d", "e")
            TestArrayUBound = UBound(arr)
        End Function
    `;

    try {
        const result = runFunc(code, 'TestArrayUBound');
        assert.strictEqual(result, 4, 'Array() with 5 elements has UBound = 4 (0-based)');
        console.log('[PASS] Array() UBound calculation');
    } catch (e: any) {
        console.log('[FAIL] Array() UBound calculation:', e.message);
    }
}

// Test 3: Array() single element
{
    const code = String.raw`
        Function TestArraySingle()
            Dim arr
            arr = Array(42)
            TestArraySingle = LBound(arr) & ":" & UBound(arr)
        End Function
    `;

    try {
        const result = runFunc(code, 'TestArraySingle');
        assert.strictEqual(result, '0:0', 'Single-element array has LBound=0, UBound=0');
        console.log('[PASS] Single-element Array()');
    } catch (e: any) {
        console.log('[FAIL] Single-element Array():', e.message);
    }
}

// Test 4: Array() empty
{
    const code = String.raw`
        Function TestArrayEmpty()
            Dim arr
            arr = Array()
            TestArrayEmpty = LBound(arr)
        End Function
    `;

    try {
        const result = runFunc(code, 'TestArrayEmpty');
        assert.strictEqual(result, 0, 'Empty Array() has LBound = 0');
        console.log('[PASS] Empty Array()');
    } catch (e: any) {
        console.log('[FAIL] Empty Array():', e.message);
    }
}

// Test 5: Array() element access by 0-based index
{
    const code = String.raw`
        Function TestArrayAccess()
            Dim arr
            arr = Array(100, 200, 300)
            TestArrayAccess = arr(0) & "," & arr(1) & "," & arr(2)
        End Function
    `;

    try {
        const result = runFunc(code, 'TestArrayAccess');
        assert.strictEqual(result, '100,200,300', 'Array() elements accessed by 0-based indexing');
        console.log('[PASS] Array() 0-based element access');
    } catch (e: any) {
        console.log('[FAIL] Array() 0-based element access:', e.message);
    }
}

// Test 6: Array() with For loop (0-based)
{
    const code = String.raw`
        Function TestArrayForLoop()
            Dim arr, i, sum
            arr = Array(1, 2, 3, 4, 5)
            sum = 0
            For i = LBound(arr) To UBound(arr)
                sum = sum + arr(i)
            Next i
            TestArrayForLoop = sum
        End Function
    `;

    try {
        const result = runFunc(code, 'TestArrayForLoop');
        assert.strictEqual(result, 15, 'For loop with 0-based Array() sums correctly');
        console.log('[PASS] For loop with Array()');
    } catch (e: any) {
        console.log('[FAIL] For loop with Array():', e.message);
    }
}

// Test 7: Array() with For Each
{
    const code = String.raw`
        Function TestArrayForEach()
            Dim arr, item, result
            arr = Array("x", "y", "z")
            result = ""
            For Each item In arr
                result = result & item
            Next item
            TestArrayForEach = result
        End Function
    `;

    try {
        const result = runFunc(code, 'TestArrayForEach');
        assert.strictEqual(result, 'xyz', 'For Each with Array() iterates all elements');
        console.log('[PASS] For Each with Array()');
    } catch (e: any) {
        console.log('[FAIL] For Each with Array():', e.message);
    }
}

// Test 8: Array() mixed types (Variant array)
{
    const code = String.raw`
        Function TestArrayMixed()
            Dim arr
            arr = Array(1, "two", 3.14, True)
            TestArrayMixed = LBound(arr) & "," & UBound(arr)
        End Function
    `;

    try {
        const result = runFunc(code, 'TestArrayMixed');
        assert.strictEqual(result, '0,3', 'Mixed-type Array() is 0-based with 4 elements');
        console.log('[PASS] Mixed-type Array()');
    } catch (e: any) {
        console.log('[FAIL] Mixed-type Array():', e.message);
    }
}

// Test 9: Array() difference from declared array bounds
{
    const code = String.raw`
        Function TestArrayVsDeclared()
            Dim declArr(1 To 3)
            Dim arrFunc
            arrFunc = Array(10, 20, 30)
            ' declArr has LBound=1, Array() has LBound=0
            TestArrayVsDeclared = LBound(declArr) & ":" & LBound(arrFunc)
        End Function
    `;

    try {
        const result = runFunc(code, 'TestArrayVsDeclared');
        assert.strictEqual(result, '1:0', 'Declared array (1-based) vs Array() function (0-based)');
        console.log('[PASS] Array() vs declared array bounds');
    } catch (e: any) {
        console.log('[FAIL] Array() vs declared array bounds:', e.message);
    }
}

// Test 10: Real-world pattern - building dynamic arrays
{
    const code = String.raw`
        Function BuildList(count)
            Dim result
            result = Array()
            For i = 1 To count
                ReDim Preserve result(LBound(result) To UBound(result) + 1)
                result(UBound(result)) = i * 10
            Next i
            BuildList = result(UBound(result))
        End Function
    `;

    try {
        const result = runFunc(code, 'BuildList', [3]);
        assert.strictEqual(result, 30, 'Building dynamic array with ReDim Preserve and Array()');
        console.log('[PASS] Dynamic array building pattern');
    } catch (e: any) {
        console.log('[FAIL] Dynamic array building pattern:', e.message);
    }
}

console.log('\n✅ Array LBound Defaults: 全テスト完了');
