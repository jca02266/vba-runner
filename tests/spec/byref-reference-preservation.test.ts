import { Lexer } from '../../src/engine/lexer';
import { Parser } from '../../src/engine/parser';
import { Evaluator } from '../../src/engine/evaluator';
import { assert } from '../../test-libs/test-runner';

function evalVBA(code: string): any {
    const tokens = new Lexer(code).tokenize();
    const ast = new Parser(tokens).parse();
    const ev = new Evaluator(console.log);
    ev.evaluate(ast);
    return ev;
}

function runFunc(code: string, name: string, args: any[] = []): any {
    return evalVBA(code).callProcedure(name, args);
}

// Test 1: Basic ByRef string reassignment
{
    const code = `
    Sub ModifyString(ByRef s As String)
        s = "Modified"
    End Sub

    Function Test1()
        Dim str As String
        str = "Original"
        ModifyString str
        Test1 = str
    End Function
    `;
    const result = runFunc(code, 'Test1');
    assert.strictEqual(result, 'Modified', 'ByRef string reassignment should affect caller');
    console.log('[PASS] Test 1: Basic ByRef string reassignment');
}

// Test 2: ByRef string concatenation
{
    const code = `
    Sub AppendToString(ByRef s As String)
        s = s & " Appended"
    End Sub

    Function Test2()
        Dim str As String
        str = "Hello"
        AppendToString str
        Test2 = str
    End Function
    `;
    const result = runFunc(code, 'Test2');
    assert.strictEqual(result, 'Hello Appended', 'ByRef string concatenation should affect caller');
    console.log('[PASS] Test 2: ByRef string concatenation');
}

// Test 3: ByRef array element modification
{
    const code = `
    Sub ModifyArrayElement(ByRef arr() As Integer)
        arr(0) = 999
    End Sub

    Function Test3()
        Dim myArr(2) As Integer
        myArr(0) = 10
        myArr(1) = 20
        myArr(2) = 30
        ModifyArrayElement myArr
        Test3 = myArr(0)
    End Function
    `;
    const result = runFunc(code, 'Test3');
    assert.strictEqual(result, 999, 'ByRef array element modification should affect caller');
    console.log('[PASS] Test 3: ByRef array element modification');
}

// Test 4: ByRef array reassignment (replacement of entire array)
{
    const code = `
    Sub ReplaceArray(ByRef arr() As Integer)
        ReDim arr(1)
        arr(0) = 100
        arr(1) = 200
    End Sub

    Function Test4()
        Dim myArr(2) As Integer
        myArr(0) = 1
        myArr(1) = 2
        myArr(2) = 3
        ReplaceArray myArr
        Test4 = myArr(0) + myArr(1)
    End Function
    `;
    const result = runFunc(code, 'Test4');
    assert.strictEqual(result, 300, 'ByRef array reassignment (ReDim) should affect caller');
    console.log('[PASS] Test 4: ByRef array reassignment with ReDim');
}

// Test 5: ByRef object property modification
{
    const code = `
    Class MyClass
        Public Value As Integer
    End Class

    Sub ModifyObject(ByRef obj As MyClass)
        obj.Value = 777
    End Sub

    Function Test5()
        Dim myObj As MyClass
        Set myObj = New MyClass
        myObj.Value = 100
        ModifyObject myObj
        Test5 = myObj.Value
    End Function
    `;
    const result = runFunc(code, 'Test5');
    assert.strictEqual(result, 777, 'ByRef object property modification should affect caller');
    console.log('[PASS] Test 5: ByRef object property modification');
}

// Test 6: ByRef object reassignment (Set with different object)
{
    const code = `
    Class MyClass
        Public Value As Integer
    End Class

    Sub ReplaceObject(ByRef obj As MyClass)
        Set obj = New MyClass
        obj.Value = 555
    End Sub

    Function Test6()
        Dim myObj As MyClass
        Set myObj = New MyClass
        myObj.Value = 100
        ReplaceObject myObj
        Test6 = myObj.Value
    End Function
    `;
    const result = runFunc(code, 'Test6');
    assert.strictEqual(result, 555, 'ByRef object reassignment (Set) should affect caller');
    console.log('[PASS] Test 6: ByRef object reassignment with Set');
}

// Test 7: ByVal vs ByRef for string - ByVal should NOT modify
{
    const code = `
    Sub ModifyStringByVal(ByVal s As String)
        s = "Modified"
    End Sub

    Function Test7()
        Dim str As String
        str = "Original"
        ModifyStringByVal str
        Test7 = str
    End Function
    `;
    const result = runFunc(code, 'Test7');
    assert.strictEqual(result, 'Original', 'ByVal string should not affect caller');
    console.log('[PASS] Test 7: ByVal string does not modify caller');
}

// Test 8: Multiple ByRef parameters - all should be affected
{
    const code = `
    Sub SwapStrings(ByRef a As String, ByRef b As String)
        Dim temp As String
        temp = a
        a = b
        b = temp
    End Sub

    Function Test8()
        Dim x As String, y As String
        x = "First"
        y = "Second"
        SwapStrings x, y
        Test8 = x & "-" & y
    End Function
    `;
    const result = runFunc(code, 'Test8');
    assert.strictEqual(result, 'Second-First', 'Multiple ByRef parameters should all affect caller');
    console.log('[PASS] Test 8: Multiple ByRef parameters');
}

// Test 9: ByRef with nested procedure calls
{
    const code = `
    Sub Increment(ByRef n As Integer)
        n = n + 1
    End Sub

    Sub DoubleIncrement(ByRef n As Integer)
        Increment n
        Increment n
    End Sub

    Function Test9()
        Dim counter As Integer
        counter = 10
        DoubleIncrement counter
        Test9 = counter
    End Function
    `;
    const result = runFunc(code, 'Test9');
    assert.strictEqual(result, 12, 'ByRef through nested procedure calls should accumulate');
    console.log('[PASS] Test 9: ByRef with nested procedure calls');
}

// Test 10: ByRef string with operations
{
    const code = `
    Sub TransformString(ByRef s As String)
        s = UCase(s)
    End Sub

    Function Test10()
        Dim str As String
        str = "hello"
        TransformString str
        Test10 = str
    End Function
    `;
    const result = runFunc(code, 'Test10');
    assert.strictEqual(result, 'HELLO', 'ByRef string with function result should work');
    console.log('[PASS] Test 10: ByRef string with function operations');
}

// Test 11: ByRef Variant - numeric
{
    const code = `
    Sub DoubleVariant(ByRef v As Variant)
        v = v * 2
    End Sub

    Function Test11()
        Dim val As Variant
        val = 25
        DoubleVariant val
        Test11 = val
    End Function
    `;
    const result = runFunc(code, 'Test11');
    assert.strictEqual(result, 50, 'ByRef Variant with numeric should work');
    console.log('[PASS] Test 11: ByRef Variant numeric');
}

// Test 12: ByRef Variant - string
{
    const code = `
    Sub ReverseVariant(ByRef v As Variant)
        v = StrReverse(v)
    End Sub

    Function Test12()
        Dim val As Variant
        val = "Hello"
        ReverseVariant val
        Test12 = val
    End Function
    `;
    const result = runFunc(code, 'Test12');
    assert.strictEqual(result, 'olleH', 'ByRef Variant with string should work');
    console.log('[PASS] Test 12: ByRef Variant string');
}

// Test 13: Array passed as ByRef - multiple modifications
{
    const code = `
    Sub ModifyArray(ByRef arr() As Integer)
        arr(0) = 111
        arr(1) = 222
        arr(2) = 333
    End Sub

    Function Test13()
        Dim myArr(2) As Integer
        myArr(0) = 1
        myArr(1) = 2
        myArr(2) = 3
        ModifyArray myArr
        Test13 = myArr(0) + myArr(1) + myArr(2)
    End Function
    `;
    const result = runFunc(code, 'Test13');
    assert.strictEqual(result, 666, 'ByRef array all elements modification should work');
    console.log('[PASS] Test 13: ByRef array multiple element modifications');
}

// Test 14: ByRef with conditional modification
{
    const code = `
    Sub ConditionalModify(ByRef s As String)
        If Len(s) > 0 Then
            s = "Modified: " & s
        End If
    End Sub

    Function Test14()
        Dim str As String
        str = "test"
        ConditionalModify str
        Test14 = str
    End Function
    `;
    const result = runFunc(code, 'Test14');
    assert.strictEqual(result, 'Modified: test', 'ByRef with conditional modification should work');
    console.log('[PASS] Test 14: ByRef with conditional modification');
}

// Test 15: VBA veteran pattern - in-place array sorting via ByRef
{
    const code = `
    Sub SimpleSort(ByRef arr() As Integer)
        Dim i As Integer, j As Integer, temp As Integer
        For i = 0 To 2
            For j = 0 To 1
                If arr(j) > arr(j + 1) Then
                    temp = arr(j)
                    arr(j) = arr(j + 1)
                    arr(j + 1) = temp
                End If
            Next j
        Next i
    End Sub

    Function Test15()
        Dim myArr(2) As Integer
        myArr(0) = 30
        myArr(1) = 10
        myArr(2) = 20
        SimpleSort myArr
        Test15 = myArr(0) + myArr(1) * 10 + myArr(2) * 100
    End Function
    `;
    const result = runFunc(code, 'Test15');
    assert.strictEqual(result, 3210, 'Common ByRef pattern: in-place array modification via sorting');
    console.log('[PASS] Test 15: VBA veteran pattern - in-place array sorting');
}

console.log('\n✅ ByRef Reference Preservation: 全テスト通過');
