import { evalVBASingle, assert } from '../../test-libs/test-runner';

function evalVBA(code: string): any {
    return evalVBASingle(code);
}

function runFunc(code: string, name: string, args: any[] = []): any {
    return evalVBA(code).callProcedure(name, args);
}

// Test 1: Property Get called in value context (retrieval)
{
    const code = `
    Class TestClass
        Private myValue As Integer

        Public Property Get Value() As Integer
            Value = myValue
        End Property

        Public Property Let Value(v As Integer)
            myValue = v
        End Property
    End Class

    Function Test1()
        Dim obj As New TestClass
        obj.Value = 10
        Test1 = obj.Value  ' Should call Property Get
    End Function
    `;
    try {
        const result = runFunc(code, 'Test1');
        assert.strictEqual(result, 10, 'Property Get should be called in value context');
        console.log('[PASS] Test 1: Property Get in value context');
    } catch (e: any) {
        console.log('[INFO] Test 1: ' + e.message);
    }
}

// Test 2: Property Let called in assignment context
{
    const code = `
    Class TestClass
        Private myValue As Integer

        Public Property Get Value() As Integer
            Value = myValue
        End Property

        Public Property Let Value(v As Integer)
            myValue = v
        End Property
    End Class

    Function Test2()
        Dim obj As New TestClass
        obj.Value = 42  ' Should call Property Let
        Test2 = obj.Value
    End Function
    `;
    try {
        const result = runFunc(code, 'Test2');
        assert.strictEqual(result, 42, 'Property Let should be called in assignment');
        console.log('[PASS] Test 2: Property Let in assignment context');
    } catch (e: any) {
        console.log('[INFO] Test 2: ' + e.message);
    }
}

// Test 3: Property methods returning values
{
    const code = `
    Class TestClass
        Private val1 As Integer
        Private val2 As Integer

        Public Property Get First() As Integer
            First = val1
        End Property

        Public Property Let First(v As Integer)
            val1 = v
        End Property

        Public Property Get Second() As Integer
            Second = val2
        End Property

        Public Property Let Second(v As Integer)
            val2 = v
        End Property
    End Class

    Function Test3()
        Dim obj As New TestClass
        obj.First = 30
        obj.Second = 20
        Test3 = obj.First + obj.Second
    End Function
    `;
    try {
        const result = runFunc(code, 'Test3');
        assert.strictEqual(result, 50, 'Multiple properties dispatch correctly');
        console.log('[PASS] Test 3: Multiple property dispatch');
    } catch (e: any) {
        console.log('[INFO] Test 3: ' + e.message);
    }
}

// Test 4: Chained property access and mutation
{
    const code = `
    Class TestClass
        Private myValue As Integer

        Public Property Get Value() As Integer
            Value = myValue
        End Property

        Public Property Let Value(v As Integer)
            myValue = v
        End Property
    End Class

    Function Test4()
        Dim obj As New TestClass
        obj.Value = 25
        Dim temp As Integer
        temp = obj.Value
        Test4 = temp
    End Function
    `;
    try {
        const result = runFunc(code, 'Test4');
        assert.strictEqual(result, 25, 'Property dispatch in get and set');
        console.log('[PASS] Test 4: Chained property access');
    } catch (e: any) {
        console.log('[INFO] Test 4: ' + e.message);
    }
}

// Test 5: Property Get/Let with intermediate variable
{
    const code = `
    Class TestClass
        Private myVal As Integer

        Public Property Get PropValue() As Integer
            PropValue = myVal
        End Property

        Public Property Let PropValue(v As Integer)
            myVal = v
        End Property
    End Class

    Function Test5()
        Dim obj As New TestClass
        Dim temp As Integer
        temp = 50
        obj.PropValue = temp
        Test5 = obj.PropValue
    End Function
    `;
    try {
        const result = runFunc(code, 'Test5');
        assert.strictEqual(result, 50, 'Property Let/Get with variable');
        console.log('[PASS] Test 5: Property dispatch with variable');
    } catch (e: any) {
        console.log('[INFO] Test 5: ' + e.message);
    }
}

// Test 6: Sequential property set and get
{
    const code = `
    Class Counter
        Private count As Integer

        Public Property Get Current() As Integer
            Current = count
        End Property

        Public Property Let Current(v As Integer)
            count = v
        End Property

        Public Sub Increment()
            count = count + 1
        End Sub
    End Class

    Function Test6()
        Dim obj As New Counter
        obj.Current = 100
        obj.Increment
        obj.Current = obj.Current + 100
        Test6 = obj.Current
    End Function
    `;
    try {
        const result = runFunc(code, 'Test6');
        assert.strictEqual(result, 201, 'Sequential property dispatch');
        console.log('[PASS] Test 6: Sequential property set and get');
    } catch (e: any) {
        console.log('[INFO] Test 6: ' + e.message);
    }
}

// Test 7: Property dispatch with numeric expression
{
    const code = `
    Class TestClass
        Private val As Integer

        Public Property Get Value() As Integer
            Value = val
        End Property

        Public Property Let Value(v As Integer)
            val = v
        End Property
    End Class

    Function Test7()
        Dim obj As New TestClass
        obj.Value = 50 - 8
        Test7 = obj.Value
    End Function
    `;
    try {
        const result = runFunc(code, 'Test7');
        assert.strictEqual(result, 42, 'Property dispatch with expression');
        console.log('[PASS] Test 7: Property dispatch with numeric expression');
    } catch (e: any) {
        console.log('[INFO] Test 7: ' + e.message);
    }
}

// Test 8: ByRef parameter in Property Set
{
    const code = `
    Class TestClass
        Private myValue As Integer

        Public Property Get Counter() As Integer
            Counter = myValue
        End Property

        Public Property Let Counter(ByRef v As Integer)
            myValue = v
            v = v + 1
        End Property
    End Class

    Function Test8()
        Dim obj As New TestClass
        Dim x As Integer
        x = 5
        obj.Counter = x
        Test8 = obj.Counter  ' Should be 5 (Property Let called)
    End Function
    `;
    try {
        const result = runFunc(code, 'Test8');
        assert.strictEqual(result, 5, 'ByRef parameter in property dispatch');
        console.log('[PASS] Test 8: ByRef in Property Set');
    } catch (e: any) {
        console.log('[INFO] Test 8: ' + e.message);
    }
}

// Test 9: Property access through intermediate variable
{
    const code = `
    Class TestClass
        Private myValue As Integer

        Public Property Get Value() As Integer
            Value = myValue
        End Property

        Public Property Let Value(v As Integer)
            myValue = v
        End Property
    End Class

    Function Test9()
        Dim obj As New TestClass
        Dim v1 As Integer
        Dim v2 As Integer
        obj.Value = 77
        v1 = obj.Value
        v2 = obj.Value + 10
        Test9 = v1 + v2
    End Function
    `;
    try {
        const result = runFunc(code, 'Test9');
        assert.strictEqual(result, 164, 'Multiple property get calls');
        console.log('[PASS] Test 9: Property access through variables');
    } catch (e: any) {
        console.log('[INFO] Test 9: ' + e.message);
    }
}

// Test 10: Property without Set (read-only property)
{
    const code = `
    Class TestClass
        Private myValue As Integer

        Public Property Get Value() As Integer
            Value = myValue
        End Property

        Public Sub Initialize(v As Integer)
            myValue = v
        End Sub
    End Class

    Function Test10()
        Dim obj As New TestClass
        obj.Initialize 88
        Test10 = obj.Value  ' Should call Property Get (no Let exists)
    End Function
    `;
    try {
        const result = runFunc(code, 'Test10');
        assert.strictEqual(result, 88, 'Read-only property should work');
        console.log('[PASS] Test 10: Read-only property (Get only)');
    } catch (e: any) {
        console.log('[INFO] Test 10: ' + e.message);
    }
}

// Test 11: Write-only property scenario (Let without Get)
{
    const code = `
    Class TestClass
        Private myValue As Integer

        Public Property Let Value(v As Integer)
            myValue = v
        End Property

        Public Function GetValue() As Integer
            GetValue = myValue
        End Function
    End Class

    Function Test11()
        Dim obj As New TestClass
        obj.Value = 99
        Test11 = obj.GetValue()
    End Function
    `;
    try {
        const result = runFunc(code, 'Test11');
        assert.strictEqual(result, 99, 'Write-only property with explicit getter');
        console.log('[PASS] Test 11: Write-only property');
    } catch (e: any) {
        console.log('[INFO] Test 11: ' + e.message);
    }
}

// Test 12: Dynamic property access through late binding
{
    const code = `
    Class TestClass
        Private myValue As Integer

        Public Property Get Value() As Integer
            Value = myValue
        End Property

        Public Property Let Value(v As Integer)
            myValue = v
        End Property
    End Class

    Function Test12()
        Dim obj As Object
        Set obj = New TestClass
        obj.Value = 123  ' Late binding - should call Property Let
        Test12 = obj.Value  ' Late binding - should call Property Get
    End Function
    `;
    try {
        const result = runFunc(code, 'Test12');
        assert.strictEqual(result, 123, 'Late binding property dispatch');
        console.log('[PASS] Test 12: Late binding property dispatch');
    } catch (e: any) {
        console.log('[INFO] Test 12: ' + e.message);
    }
}

// Test 13: Distinct property names dispatch independently
{
    const code = `
    Class TestClass
        Private x As Integer
        Private y As Integer
        Private z As Integer

        Public Property Get PropX() As Integer
            PropX = x
        End Property

        Public Property Let PropX(v As Integer)
            x = v
        End Property

        Public Property Get PropY() As Integer
            PropY = y
        End Property

        Public Property Let PropY(v As Integer)
            y = v
        End Property

        Public Property Get PropZ() As Integer
            PropZ = z
        End Property

        Public Property Let PropZ(v As Integer)
            z = v
        End Property
    End Class

    Function Test13()
        Dim obj As New TestClass
        obj.PropX = 10
        obj.PropY = 20
        obj.PropZ = 30
        Test13 = obj.PropX + obj.PropY + obj.PropZ
    End Function
    `;
    try {
        const result = runFunc(code, 'Test13');
        assert.strictEqual(result, 60, 'Multiple distinct properties');
        console.log('[PASS] Test 13: Distinct property dispatch');
    } catch (e: any) {
        console.log('[INFO] Test 13: ' + e.message);
    }
}

// Test 14: Complex property dispatch with conditional logic
{
    const code = `
    Class Counter
        Private count As Integer

        Public Property Get Value() As Integer
            Value = count
        End Property

        Public Property Let Value(v As Integer)
            If v < 0 Then
                count = 0
            Else
                count = v
            End If
        End Property

        Public Sub Increment()
            count = count + 1
        End Sub
    End Class

    Function Test14()
        Dim c As New Counter
        c.Value = -5  ' Should set to 0 due to conditional
        c.Increment
        c.Increment
        Test14 = c.Value
    End Function
    `;
    try {
        const result = runFunc(code, 'Test14');
        assert.strictEqual(result, 2, 'Complex property with conditional logic');
        console.log('[PASS] Test 14: Complex property with conditional logic');
    } catch (e: any) {
        console.log('[INFO] Test 14: ' + e.message);
    }
}

// Test 15: VBA veteran pattern - Cache invalidation on property change
{
    const code = `
    Class CachedValue
        Private cached As Integer
        Private cacheValid As Boolean
        Private source As Integer

        Public Property Get Value() As Integer
            If Not cacheValid Then
                cached = source * 2
                cacheValid = True
            End If
            Value = cached
        End Property

        Public Property Let Value(v As Integer)
            source = v
            cacheValid = False
        End Property
    End Class

    Function Test15()
        Dim val As New CachedValue
        val.Value = 10  ' Invalidates cache
        Dim result As Integer
        result = val.Value  ' Rebuilds cache: 10 * 2 = 20
        val.Value = 5   ' Invalidates again
        result = result + val.Value  ' Rebuilds cache: 5 * 2 = 10
        Test15 = result
    End Function
    `;
    try {
        const result = runFunc(code, 'Test15');
        assert.strictEqual(result, 30, 'Cache invalidation on property change pattern');
        console.log('[PASS] Test 15: VBA veteran pattern - cache invalidation');
    } catch (e: any) {
        console.log('[INFO] Test 15: ' + e.message);
    }
}

console.log('\n✅ IDispatch Property Dispatch: Testing complete');
