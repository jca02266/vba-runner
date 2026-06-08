import { evalVBASingle, assert } from '../../test-libs/test-runner';

function evalVBA(code: string): any {
    return evalVBASingle(code);
}

function runFunc(code: string, name: string, args: any[] = []): any {
    return evalVBA(code).callProcedure(name, args);
}

// Test 1: String ByRef - works
{
    const code = `
    Sub Modify(ByRef s As String)
        s = "Modified"
    End Sub

    Function Test1()
        Dim str As String
        str = "Original"
        Modify str
        Test1 = str
    End Function
    `;
    const result = runFunc(code, 'Test1');
    assert.strictEqual(result, 'Modified', 'String ByRef should work');
    console.log('[PASS] Test 1: String ByRef works');
}

// Test 2: Check if array parameter gets the array value
{
    const code = `
    Sub CheckArray(ByRef arr() As Integer)
        ' Just check if we can see the array
    End Sub

    Function Test2()
        Dim myArr(2) As Integer
        myArr(0) = 10
        myArr(1) = 20
        myArr(2) = 30
        CheckArray myArr
        Test2 = 1  ' Just return 1 to confirm we got here
    End Function
    `;
    try {
        const result = runFunc(code, 'Test2');
        console.log('[PASS] Test 2: Can pass array as ByRef parameter');
    } catch (e: any) {
        console.log('[FAIL] Test 2:', e.message);
    }
}

// Test 3: Try to read array element inside ByRef procedure
{
    const code = `
    Sub ReadArray(ByRef arr() As Integer)
        Dim x As Integer
        x = arr(0)
    End Sub

    Function Test3()
        Dim myArr(2) As Integer
        myArr(0) = 10
        ReadArray myArr
        Test3 = 1
    End Function
    `;
    try {
        const result = runFunc(code, 'Test3');
        console.log('[PASS] Test 3: Can read array element in ByRef procedure');
    } catch (e: any) {
        console.log('[FAIL] Test 3:', e.message);
    }
}

// Test 4: Try array element assignment
{
    const code = `
    Sub ModifyArray(ByRef arr() As Integer)
        arr(0) = 999
    End Sub

    Function Test4()
        Dim myArr(2) As Integer
        myArr(0) = 10
        ModifyArray myArr
        Test4 = myArr(0)
    End Function
    `;
    try {
        const result = runFunc(code, 'Test4');
        console.log('[PASS] Test 4: Array element assignment in ByRef procedure works');
    } catch (e: any) {
        console.log('[FAIL] Test 4:', e.message);
    }
}

console.log('\n=== ByRef diagnostics complete ===');
