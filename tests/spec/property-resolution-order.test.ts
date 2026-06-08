import { evalVBASingle, assert } from '../../test-libs/test-runner';

function evalVBA(code: string): any {
    return evalVBASingle(code);
}

function runFunc(code: string, name: string, args: any[] = []): any {
    return evalVBA(code).callProcedure(name, args);
}

// Test 1: Property Get called on read
{
    const code = `
    Class MyClass
        Private pValue As Integer

        Public Property Get Value() As Integer
            Value = pValue
        End Property

        Public Property Let Value(v As Integer)
            pValue = v
        End Property
    End Class

    Function Test1()
        Dim obj As MyClass
        Set obj = New MyClass
        obj.Value = 42
        Test1 = obj.Value
    End Function
    `;
    const result = runFunc(code, 'Test1');
    assert.strictEqual(result, 42, 'Property Get should be called on read');
    console.log('[PASS] Test 1: Property Get called on read');
}

// Test 2: Property Let called on value assignment
{
    const code = `
    Class MyClass
        Private pValue As Integer

        Public Property Get Value() As Integer
            Value = pValue
        End Property

        Public Property Let Value(v As Integer)
            pValue = v * 2
        End Property
    End Class

    Function Test2()
        Dim obj As MyClass
        Set obj = New MyClass
        obj.Value = 20
        Test2 = obj.Value
    End Function
    `;
    const result = runFunc(code, 'Test2');
    assert.strictEqual(result, 40, 'Property Let should be called on value assignment');
    console.log('[PASS] Test 2: Property Let called on value assignment');
}

// Test 3: Property with only Get (read-only via property)
{
    const code = `
    Class MyClass
        Private pValue As Integer

        Public Property Get Value() As Integer
            Value = pValue
        End Property

        Public Sub SetValue(v As Integer)
            pValue = v
        End Sub
    End Class

    Function Test3()
        Dim obj As MyClass
        Set obj = New MyClass
        obj.SetValue 100
        Test3 = obj.Value
    End Function
    `;
    const result = runFunc(code, 'Test3');
    assert.strictEqual(result, 100, 'Property Get only (read-only property)');
    console.log('[PASS] Test 3: Property Get only pattern');
}

// Test 4: Property Get precedence in expression context
{
    const code = `
    Class MyClass
        Public Property Get Value() As Integer
            Value = 77
        End Property

        Public Property Let Value(v As Integer)
            ' Not executed in this test
        End Property
    End Class

    Function Test4()
        Dim obj As MyClass
        Set obj = New MyClass
        Dim result As Integer
        result = obj.Value + 3
        Test4 = result
    End Function
    `;
    const result = runFunc(code, 'Test4');
    assert.strictEqual(result, 80, 'Property Get in expression context');
    console.log('[PASS] Test 4: Property Get in expression context');
}

// Test 5: Multiple Property Let with same name (should have different parameter lists)
{
    const code = `
    Class MyClass
        Private pValue As String

        Public Property Get Value() As String
            Value = pValue
        End Property

        Public Property Let Value(v As String)
            pValue = v
        End Property
    End Class

    Function Test5()
        Dim obj As MyClass
        Set obj = New MyClass
        obj.Value = "Hello"
        Test5 = obj.Value
    End Function
    `;
    const result = runFunc(code, 'Test5');
    assert.strictEqual(result, 'Hello', 'Property with same name across Get/Let');
    console.log('[PASS] Test 5: Property with same name across Get/Let');
}

// Test 6: Property Let called on string assignment
{
    const code = `
    Class MyClass
        Private pValue As String

        Public Property Get Value() As String
            Value = pValue
        End Property

        Public Property Let Value(v As String)
            pValue = LCase(v)
        End Property
    End Class

    Function Test6()
        Dim obj As MyClass
        Set obj = New MyClass
        obj.Value = "HELLO"
        Test6 = obj.Value
    End Function
    `;
    const result = runFunc(code, 'Test6');
    assert.strictEqual(result, 'hello', 'Property Let with transformation');
    console.log('[PASS] Test 6: Property Let with transformation');
}

// Test 7: Property Let with side effect
{
    const code = `
    Class MyClass
        Private pValue As Integer
        Private pSideEffect As Integer

        Public Property Get Value() As Integer
            Value = pValue
        End Property

        Public Property Let Value(v As Integer)
            pValue = v
            pSideEffect = v + 1000
        End Property

        Public Function GetSideEffect() As Integer
            GetSideEffect = pSideEffect
        End Function
    End Class

    Function Test7()
        Dim obj As MyClass
        Set obj = New MyClass
        obj.Value = 25
        Test7 = obj.GetSideEffect()
    End Function
    `;
    const result = runFunc(code, 'Test7');
    assert.strictEqual(result, 1025, 'Property Let side effect executes');
    console.log('[PASS] Test 7: Property Let side effect');
}

// Test 8: Property Get without Let (read-only property)
{
    const code = `
    Class MyClass
        Private pValue As Integer

        Public Sub Initialize(v As Integer)
            pValue = v
        End Sub

        Public Property Get Value() As Integer
            Value = pValue
        End Property
    End Class

    Function Test8()
        Dim obj As MyClass
        Set obj = New MyClass
        obj.Initialize 88
        Test8 = obj.Value
    End Function
    `;
    const result = runFunc(code, 'Test8');
    assert.strictEqual(result, 88, 'Read-only property (Get without Let)');
    console.log('[PASS] Test 8: Read-only property');
}

// Test 9: Property returning numeric value in calculation
{
    const code = `
    Class MyClass
        Private pValue As Integer

        Public Property Get Value() As Integer
            Value = pValue
        End Property

        Public Property Let Value(v As Integer)
            pValue = v
        End Property
    End Class

    Function Test9()
        Dim obj As MyClass
        Set obj = New MyClass
        obj.Value = 10
        Dim result As Integer
        result = (obj.Value * 2) + 5
        Test9 = result
    End Function
    `;
    const result = runFunc(code, 'Test9');
    assert.strictEqual(result, 25, 'Property in numeric calculation');
    console.log('[PASS] Test 9: Property in numeric calculation');
}

// Test 10: Property Let with integer conversion
{
    const code = `
    Class MyClass
        Private pValue As Integer

        Public Property Get Value() As Integer
            Value = pValue
        End Property

        Public Property Let Value(v As Integer)
            pValue = v
        End Property
    End Class

    Function Test10()
        Dim obj As MyClass
        Set obj = New MyClass
        obj.Value = CLng(42)
        Test10 = obj.Value
    End Function
    `;
    const result = runFunc(code, 'Test10');
    assert.strictEqual(result, 42, 'Property Let with integer value');
    console.log('[PASS] Test 10: Property Let with integer value');
}

// Test 11: Property in conditional
{
    const code = `
    Class MyClass
        Public Property Get IsValid() As Boolean
            IsValid = True
        End Property
    End Class

    Function Test11()
        Dim obj As MyClass
        Set obj = New MyClass
        If obj.IsValid Then
            Test11 = 1
        Else
            Test11 = 0
        End If
    End Function
    `;
    const result = runFunc(code, 'Test11');
    assert.strictEqual(result, 1, 'Property in conditional');
    console.log('[PASS] Test 11: Property in conditional');
}

// Test 12: Property in loop
{
    const code = `
    Class MyClass
        Private pValue As Integer

        Public Property Get Value() As Integer
            Value = pValue
        End Property

        Public Property Let Value(v As Integer)
            pValue = v
        End Property
    End Class

    Function Test12()
        Dim obj As MyClass
        Set obj = New MyClass
        Dim total As Integer
        total = 0
        Dim i As Integer
        For i = 1 To 5
            obj.Value = i
            total = total + obj.Value
        Next i
        Test12 = total
    End Function
    `;
    const result = runFunc(code, 'Test12');
    assert.strictEqual(result, 15, 'Property in loop context');
    console.log('[PASS] Test 12: Property in loop');
}

// Test 13: Multiple properties in same class
{
    const code = `
    Class MyClass
        Private pX As Integer
        Private pY As Integer

        Public Property Get X() As Integer
            X = pX
        End Property

        Public Property Let X(v As Integer)
            pX = v
        End Property

        Public Property Get Y() As Integer
            Y = pY
        End Property

        Public Property Let Y(v As Integer)
            pY = v
        End Property
    End Class

    Function Test13()
        Dim obj As MyClass
        Set obj = New MyClass
        obj.X = 30
        obj.Y = 40
        Test13 = obj.X + obj.Y
    End Function
    `;
    const result = runFunc(code, 'Test13');
    assert.strictEqual(result, 70, 'Multiple independent properties');
    console.log('[PASS] Test 13: Multiple properties');
}

// Test 14: Chained property Let
{
    const code = `
    Class MyClass
        Private pValue As Integer

        Public Property Get Value() As Integer
            Value = pValue
        End Property

        Public Property Let Value(v As Integer)
            pValue = v
        End Property
    End Class

    Function Test14()
        Dim a As MyClass, b As MyClass
        Set a = New MyClass
        Set b = New MyClass
        a.Value = 50
        b.Value = a.Value + 10
        Test14 = b.Value
    End Function
    `;
    const result = runFunc(code, 'Test14');
    assert.strictEqual(result, 60, 'Chained property access in assignment');
    console.log('[PASS] Test 14: Chained property access in assignment');
}

// Test 15: VBA veteran pattern - property-based data binding
{
    const code = `
    Class DataModel
        Private pData As String

        Public Property Get Data() As String
            Data = pData
        End Property

        Public Property Let Data(v As String)
            pData = UCase(v)
        End Property
    End Class

    Function Test15()
        Dim model As DataModel
        Set model = New DataModel
        model.Data = "hello"
        Test15 = model.Data
    End Function
    `;
    const result = runFunc(code, 'Test15');
    assert.strictEqual(result, 'HELLO', 'Common pattern: property-based data binding with transformation');
    console.log('[PASS] Test 15: VBA veteran pattern - data binding');
}

console.log('\n✅ Property Get/Let/Set Resolution: 全テスト通過');
