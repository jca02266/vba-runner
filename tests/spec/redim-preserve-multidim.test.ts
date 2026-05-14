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

// Test 1: 2D array - ReDim Preserve can change last dimension upper bound
{
    const code = String.raw`
        Function Test2DPreserveLast()
            Dim arr(1 To 3, 1 To 2)
            arr(1, 1) = 100
            arr(1, 2) = 200
            ReDim Preserve arr(1 To 3, 1 To 5)
            Test2DPreserveLast = arr(1, 1) & "," & arr(1, 2) & "," & UBound(arr, 2)
        End Function
    `;

    try {
        const result = runFunc(code, 'Test2DPreserveLast');
        assert.strictEqual(result, '100,200,5', '2D ReDim Preserve can expand last dimension');
        console.log('[PASS] 2D array - expand last dimension');
    } catch (e: any) {
        console.log('[FAIL] 2D array - expand last dimension:', e.message);
    }
}

// Test 2: 2D array - ReDim Preserve can shrink last dimension upper bound
{
    const code = String.raw`
        Function Test2DPreserveShrink()
            Dim arr(1 To 3, 1 To 5)
            arr(1, 1) = 100
            arr(1, 3) = 300
            ReDim Preserve arr(1 To 3, 1 To 2)
            Test2DPreserveShrink = arr(1, 1) & "," & UBound(arr, 2)
        End Function
    `;

    try {
        const result = runFunc(code, 'Test2DPreserveShrink');
        assert.strictEqual(result, '100,2', '2D ReDim Preserve can shrink last dimension');
        console.log('[PASS] 2D array - shrink last dimension');
    } catch (e: any) {
        console.log('[FAIL] 2D array - shrink last dimension:', e.message);
    }
}

// Test 3: 2D array - ReDim Preserve CANNOT change first dimension (Error 9)
{
    const code = String.raw`
        Function Test2DPreserveFirstDim()
            Dim arr(1 To 3, 1 To 2)
            On Error GoTo ErrHandler
            ReDim Preserve arr(1 To 5, 1 To 2)
            Test2DPreserveFirstDim = "ERROR_NOT_RAISED"
            Exit Function
        ErrHandler:
            Test2DPreserveFirstDim = Err.Number
        End Function
    `;

    try {
        const result = runFunc(code, 'Test2DPreserveFirstDim');
        assert.strictEqual(result, 9, 'ReDim Preserve cannot change first dimension - Error 9');
        console.log('[PASS] 2D array - cannot change first dimension');
    } catch (e: any) {
        console.log('[FAIL] 2D array - cannot change first dimension:', e.message);
    }
}

// Test 4: 2D array - ReDim Preserve CANNOT change lower bound of first dimension (Error 9)
{
    const code = String.raw`
        Function Test2DPreserveLowerBound()
            Dim arr(1 To 3, 1 To 2)
            On Error GoTo ErrHandler
            ReDim Preserve arr(0 To 3, 1 To 2)
            Test2DPreserveLowerBound = "ERROR_NOT_RAISED"
            Exit Function
        ErrHandler:
            Test2DPreserveLowerBound = Err.Number
        End Function
    `;

    try {
        const result = runFunc(code, 'Test2DPreserveLowerBound');
        assert.strictEqual(result, 9, 'ReDim Preserve cannot change lower bound - Error 9');
        console.log('[PASS] 2D array - cannot change lower bound');
    } catch (e: any) {
        console.log('[FAIL] 2D array - cannot change lower bound:', e.message);
    }
}

// Test 5: 2D array - ReDim Preserve CANNOT change lower bound of last dimension (Error 9)
{
    const code = String.raw`
        Function Test2DPreserveLastLowerBound()
            Dim arr(1 To 3, 1 To 2)
            On Error GoTo ErrHandler
            ReDim Preserve arr(1 To 3, 0 To 2)
            Test2DPreserveLastLowerBound = "ERROR_NOT_RAISED"
            Exit Function
        ErrHandler:
            Test2DPreserveLastLowerBound = Err.Number
        End Function
    `;

    try {
        const result = runFunc(code, 'Test2DPreserveLastLowerBound');
        assert.strictEqual(result, 9, 'ReDim Preserve cannot change lower bound of last dimension - Error 9');
        console.log('[PASS] 2D array - cannot change lower bound of last dimension');
    } catch (e: any) {
        console.log('[FAIL] 2D array - cannot change lower bound of last dimension:', e.message);
    }
}

// Test 6: 3D array - ReDim Preserve can only change last dimension
{
    const code = String.raw`
        Function Test3DPreserveLast()
            Dim arr(1 To 2, 1 To 3, 1 To 2)
            arr(1, 1, 1) = 999
            ReDim Preserve arr(1 To 2, 1 To 3, 1 To 5)
            Test3DPreserveLast = arr(1, 1, 1) & "," & UBound(arr, 3)
        End Function
    `;

    try {
        const result = runFunc(code, 'Test3DPreserveLast');
        assert.strictEqual(result, '999,5', '3D ReDim Preserve can expand last dimension');
        console.log('[PASS] 3D array - expand last dimension only');
    } catch (e: any) {
        console.log('[FAIL] 3D array - expand last dimension only:', e.message);
    }
}

// Test 7: 3D array - ReDim Preserve CANNOT change middle dimension (Error 9)
{
    const code = String.raw`
        Function Test3DPreserveMiddleDim()
            Dim arr(1 To 2, 1 To 3, 1 To 2)
            On Error GoTo ErrHandler
            ReDim Preserve arr(1 To 2, 1 To 5, 1 To 2)
            Test3DPreserveMiddleDim = "ERROR_NOT_RAISED"
            Exit Function
        ErrHandler:
            Test3DPreserveMiddleDim = Err.Number
        End Function
    `;

    try {
        const result = runFunc(code, 'Test3DPreserveMiddleDim');
        assert.strictEqual(result, 9, 'ReDim Preserve cannot change middle dimension - Error 9');
        console.log('[PASS] 3D array - cannot change middle dimension');
    } catch (e: any) {
        console.log('[FAIL] 3D array - cannot change middle dimension:', e.message);
    }
}

// Test 8: 1D array - ReDim Preserve can change upper bound (special case)
{
    const code = String.raw`
        Function Test1DPreserve()
            Dim arr(1 To 5)
            arr(1) = 100
            arr(3) = 300
            ReDim Preserve arr(1 To 10)
            Test1DPreserve = arr(1) & "," & arr(3) & "," & UBound(arr)
        End Function
    `;

    try {
        const result = runFunc(code, 'Test1DPreserve');
        assert.strictEqual(result, '100,300,10', '1D ReDim Preserve works normally');
        console.log('[PASS] 1D array - ReDim Preserve expands normally');
    } catch (e: any) {
        console.log('[FAIL] 1D array - ReDim Preserve expands normally:', e.message);
    }
}

// Test 9: ReDim Preserve without keyword allows all changes
{
    const code = String.raw`
        Function TestRedimWithoutPreserve()
            Dim arr(1 To 3, 1 To 2)
            arr(1, 1) = 100
            ReDim arr(1 To 5, 1 To 5)
            ' Data should be reset, only new dimensions matter
            TestRedimWithoutPreserve = UBound(arr, 1) & "," & UBound(arr, 2)
        End Function
    `;

    try {
        const result = runFunc(code, 'TestRedimWithoutPreserve');
        assert.strictEqual(result, '5,5', 'ReDim without Preserve resets all data and allows dimension changes');
        console.log('[PASS] ReDim without Preserve allows full resizing');
    } catch (e: any) {
        console.log('[FAIL] ReDim without Preserve allows full resizing:', e.message);
    }
}

// Test 10: ReDim Preserve with 2D array - real-world pattern (expanding last dim for appending)
{
    const code = String.raw`
        Function TestRealWorldAppend()
            Dim data(1 To 3, 1 To 2)
            data(1, 1) = "A"
            data(1, 2) = "B"
            data(2, 1) = "C"
            data(2, 2) = "D"
            ' Expand last dimension to add more columns
            ReDim Preserve data(1 To 3, 1 To 4)
            data(1, 3) = "E"
            data(1, 4) = "F"
            TestRealWorldAppend = data(1, 1) & data(1, 2) & data(1, 3) & data(1, 4)
        End Function
    `;

    try {
        const result = runFunc(code, 'TestRealWorldAppend');
        assert.strictEqual(result, 'ABEF', 'Real-world pattern: expand last dimension to append columns');
        console.log('[PASS] Real-world pattern - expanding columns with ReDim Preserve');
    } catch (e: any) {
        console.log('[FAIL] Real-world pattern - expanding columns with ReDim Preserve:', e.message);
    }
}

console.log('\n✅ ReDim Preserve Multi-Dimensional: 全テスト完了');
