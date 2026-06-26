import { evalVBASingle, assert } from '../../test-libs/test-runner';

function evalVBA(code: string): any {
    return evalVBASingle(code);
}

function runFunc(code: string, name: string, args: any[] = []): any {
    return evalVBA(code).callProcedure(name, args);
}

// Test 1: Basic Class_Terminate invocation on Set = Nothing
{
    const code = `
    Dim terminateCount As Integer

    Class MyClass
        Public Sub Class_Terminate()
            terminateCount = terminateCount + 1
        End Sub
    End Class

    Function Test1()
        Dim obj As MyClass
        Set obj = New MyClass
        Set obj = Nothing
        Test1 = terminateCount
    End Function
    `;
    const result = runFunc(code, 'Test1');
    assert.strictEqual(result, 1, 'Class_Terminate should be called once when object is set to Nothing');
    console.log('[PASS] Test 1: Basic Class_Terminate invocation');
}

// Test 2: Two objects with mutual references - one set to Nothing
{
    const code = `
    Dim terminateCountA As Integer
    Dim terminateCountB As Integer

    Class ClassA
        Public RefB As Object
        Public Sub Class_Terminate()
            terminateCountA = terminateCountA + 1
            Set RefB = Nothing
        End Sub
    End Class

    Class ClassB
        Public RefA As Object
        Public Sub Class_Terminate()
            terminateCountB = terminateCountB + 1
        End Sub
    End Class

    Function Test2()
        Dim objA As ClassA
        Dim objB As ClassB
        Set objA = New ClassA
        Set objB = New ClassB
        Set objA.RefB = objB
        Set objB.RefA = objA
        Set objA = Nothing
        Test2 = terminateCountA + terminateCountB * 10
    End Function
    `;
    const result = runFunc(code, 'Test2');
    // When objA is set to Nothing, Class_Terminate for A is called, which sets RefB to Nothing
    // This should trigger B's termination as well
    assert.strictEqual(result, 11, 'Both Class_Terminate handlers should execute (A once, B once)');
    console.log('[PASS] Test 2: Mutual references with Set = Nothing');
}

// Test 3: Three objects in a chain - breaking the chain
{
    const code = `
    Dim chain1 As Integer
    Dim chain2 As Integer
    Dim chain3 As Integer

    Class ChainObj
        Public LinkedNode As Object
        Public ID As Integer
        Public Sub Class_Terminate()
            If ID = 1 Then
                chain1 = chain1 + 1
            ElseIf ID = 2 Then
                chain2 = chain2 + 1
            ElseIf ID = 3 Then
                chain3 = chain3 + 1
            End If
            Set LinkedNode = Nothing
        End Sub
    End Class

    Function Test3()
        Dim obj1 As ChainObj
        Dim obj2 As ChainObj
        Dim obj3 As ChainObj
        Set obj1 = New ChainObj
        Set obj2 = New ChainObj
        Set obj3 = New ChainObj
        obj1.ID = 1
        obj2.ID = 2
        obj3.ID = 3
        Set obj1.LinkedNode = obj2
        Set obj2.LinkedNode = obj3
        Set obj1 = Nothing
        Set obj2 = Nothing
        Set obj3 = Nothing
        Test3 = chain1 + chain2 * 10 + chain3 * 100
    End Function
    `;
    const result = runFunc(code, 'Test3');
    // All three should terminate
    assert.strictEqual(result, 111, 'All three objects should call Class_Terminate');
    console.log('[PASS] Test 3: Chain of three objects termination');
}

// Test 4: Class_Terminate at most once
{
    const code = `
    Dim terminateAttempts As Integer

    Class ReusableObj
        Public MakeAccessible As Boolean
        Public Sub Class_Terminate()
            terminateAttempts = terminateAttempts + 1
        End Sub
    End Class

    Function Test4()
        Dim obj As ReusableObj
        Set obj = New ReusableObj
        Set obj = Nothing
        Test4 = terminateAttempts
    End Function
    `;
    const result = runFunc(code, 'Test4');
    assert.strictEqual(result, 1, 'Class_Terminate should execute at most once');
    console.log('[PASS] Test 4: Class_Terminate executes at most once');
}

// Test 5: Multiple independent objects
{
    const code = `
    Dim count As Integer

    Class SimpleObj
        Public Sub Class_Terminate()
            count = count + 1
        End Sub
    End Class

    Function Test5()
        Dim obj1 As SimpleObj
        Dim obj2 As SimpleObj
        Dim obj3 As SimpleObj
        Set obj1 = New SimpleObj
        Set obj2 = New SimpleObj
        Set obj3 = New SimpleObj
        Set obj1 = Nothing
        Set obj2 = Nothing
        Set obj3 = Nothing
        Test5 = count
    End Function
    `;
    const result = runFunc(code, 'Test5');
    assert.strictEqual(result, 3, 'Three independent objects should each call Class_Terminate once');
    console.log('[PASS] Test 5: Multiple independent objects termination');
}

// Test 6: Class_Terminate with cleanup logic
{
    const code = `
    Dim cleanupLog As String

    Class Resource
        Public Name As String
        Public Sub Class_Terminate()
            cleanupLog = cleanupLog & "cleanup:" & Name & ";"
        End Sub
    End Class

    Function Test6()
        Dim res1 As Resource
        Dim res2 As Resource
        Set res1 = New Resource
        res1.Name = "File1"
        Set res2 = New Resource
        res2.Name = "File2"
        Set res1 = Nothing
        Set res2 = Nothing
        Test6 = Len(cleanupLog) > 0
    End Function
    `;
    const result = runFunc(code, 'Test6');
    assert.strictEqual(result, -1, 'Class_Terminate should execute cleanup logic');
    console.log('[PASS] Test 6: Class_Terminate with cleanup logic');
}

// Test 7: Circular reference with self-reference
{
    const code = `
    Dim terminateCount As Integer

    Class SelfRef
        Public SelfLink As Object
        Public Sub Class_Terminate()
            terminateCount = terminateCount + 1
            Set SelfLink = Nothing
        End Sub
    End Class

    Function Test7()
        Dim obj As SelfRef
        Set obj = New SelfRef
        Set obj.SelfLink = obj
        Set obj = Nothing
        Test7 = terminateCount
    End Function
    `;
    const result = runFunc(code, 'Test7');
    assert.strictEqual(result, 1, 'Self-referencing object should terminate once');
    console.log('[PASS] Test 7: Self-referencing object');
}

// Test 8: VBA veteran pattern - resource management with circular refs
{
    const code = `
    Dim openCount As Integer
    Dim closeCount As Integer

    Class File
        Public Buffer As Object
        Public Sub Open()
            openCount = openCount + 1
        End Sub
        Public Sub Class_Terminate()
            closeCount = closeCount + 1
            Set Buffer = Nothing
        End Sub
    End Class

    Class Buffer
        Public Parent As Object
        Public Data As String
    End Class

    Function Test8()
        Dim f As File
        Dim b As Buffer
        Set f = New File
        Set b = New Buffer
        f.Data = "test"
        f.Open()
        Set f.Buffer = b
        Set b.Parent = f
        Set f = Nothing
        Set b = Nothing
        Test8 = openCount * 10 + closeCount
    End Function
    `;
    const result = runFunc(code, 'Test8');
    assert.strictEqual(result, 11, 'File should be opened and closed properly with circular refs');
    console.log('[PASS] Test 8: Resource management with circular references');
}

// Test 9: Termination order with dependent objects
{
    const code = `
    Dim terminationOrder As String

    Class Parent
        Public Child As Object
        Public Sub Class_Terminate()
            terminationOrder = terminationOrder & "P"
            Set Child = Nothing
        End Sub
    End Class

    Class Child
        Public Sub Class_Terminate()
            terminationOrder = terminationOrder & "C"
        End Sub
    End Class

    Function Test9()
        Dim p As Parent
        Dim c As Child
        Set p = New Parent
        Set c = New Child
        Set p.Child = c
        Set p = Nothing
        Set c = Nothing
        Test9 = Len(terminationOrder) = 2
    End Function
    `;
    const result = runFunc(code, 'Test9');
    assert.strictEqual(result, -1, 'Both Parent and Child should terminate');
    console.log('[PASS] Test 9: Termination order with dependent objects');
}

// Test 10: Nested circular references with explicit cleanup
{
    const code = `
    Dim globalCount As Integer

    Class Scoped
        Public Peer As Object
        Public Sub Class_Terminate()
            globalCount = globalCount + 1
            Set Peer = Nothing
        End Sub
    End Class

    Function Test10()
        Dim s1 As Scoped
        Dim s2 As Scoped
        Set s1 = New Scoped
        Set s2 = New Scoped
        Set s1.Peer = s2
        Set s2.Peer = s1
        Set s1 = Nothing
        Set s2 = Nothing
        Test10 = globalCount
    End Function
    `;
    const result = runFunc(code, 'Test10');
    assert.strictEqual(result, 2, 'Both objects should terminate with explicit cleanup');
    console.log('[PASS] Test 10: Nested circular references with explicit cleanup');
}

// Test 11: B-3 — オブジェクトを Dictionary に格納後、変数を別インスタンスに切り替えても
//             Dictionary 内のオブジェクトの Class_Terminate は呼ばれない
{
    const code = `
    Dim terminateLog As String

    Class Tracked
        Public Name As String
        Private Sub Class_Terminate()
            terminateLog = terminateLog & Name & " terminated;"
        End Sub
    End Class

    Function Test11() As String
        Dim dict As Object
        Set dict = CreateObject("Scripting.Dictionary")
        terminateLog = ""

        Dim p As Tracked
        Set p = New Tracked
        p.Name = "A"
        dict.Add "a", p

        ' 別インスタンスに切り替え: dict 内の "A" の Terminate は呼ばれない
        Set p = New Tracked
        p.Name = "B"
        dict.Add "b", p

        ' この時点では terminateLog は空のはず
        Dim midLog As String
        midLog = terminateLog

        ' 明示的に Nothing にすることで "B" の Terminate が呼ばれる
        Set p = Nothing

        Test11 = midLog & "|" & terminateLog
    End Function
    `;
    const result = runFunc(code, 'Test11');
    assert.strictEqual(result, '|B terminated;', 'B-3: Dictionary 内オブジェクトの早期 Terminate が起きない');
    console.log('[PASS] B-3: Dictionary に格納済みオブジェクトの早期 Terminate が発生しない');
}

// Test 12: Scope exit triggers Class_Terminate (no explicit Set = Nothing needed)
{
    const code = `
    Dim count As Integer

    Class AutoClose
        Public Sub Class_Terminate()
            count = count + 1
        End Sub
    End Class

    Sub CreateAndForget()
        Dim obj As AutoClose
        Set obj = New AutoClose
        ' No explicit Set obj = Nothing — scope exit should trigger Terminate
    End Sub

    Function GetCount() As Integer
        GetCount = count
    End Function
    `;
    const runner = evalVBA(code);
    runner.callProcedure('CreateAndForget', []);
    const result = runner.callProcedure('GetCount', []);
    assert.strictEqual(result, 1, 'Scope exit should trigger Class_Terminate for local VBA objects');
    console.log('[PASS] Test 12: Scope exit triggers Class_Terminate');
}

// Test 13: Scope exit does NOT terminate params (caller still holds them)
{
    const code = `
    Dim count As Integer

    Class BorrowedObj
        Public Sub Class_Terminate()
            count = count + 1
        End Sub
    End Class

    Sub BorrowObject(obj As BorrowedObj)
        ' param — scope exit must NOT terminate
    End Sub

    Function Test13() As Integer
        Dim obj As BorrowedObj
        Set obj = New BorrowedObj
        BorrowObject obj
        Test13 = count  ' must be 0 — obj still alive
        Set obj = Nothing
        ' count = 1 here, but Test13 already captured 0
    End Function
    `;
    const result = runFunc(code, 'Test13');
    assert.strictEqual(result, 0, 'Params must not be terminated at callee scope exit');
    console.log('[PASS] Test 13: Params are not terminated at callee scope exit');
}

// Test 14: Two local vars pointing to same instance — Terminate fires exactly once
{
    const code = `
    Dim count As Integer

    Class Counted
        Public Sub Class_Terminate()
            count = count + 1
        End Sub
    End Class

    Sub Test14Sub()
        Dim a As Counted
        Dim b As Counted
        Set a = New Counted
        Set b = a   ' same instance
        ' scope exit: a → Terminate (count=1), b → already terminated, skip
    End Sub

    Function Test14() As Integer
        Test14Sub
        Test14 = count
    End Function
    `;
    const result = runFunc(code, 'Test14');
    assert.strictEqual(result, 1, 'Two local vars to same instance should Terminate exactly once');
    console.log('[PASS] Test 14: Same instance via two local vars terminates exactly once');
}

// --- 仕様バグ修正: Dim w As New T を一度もアクセスせず Set w = Nothing すると
// Class_Initialize/Terminate が呼ばれなかった ---
{
    const code = `
    Dim log As String
    Class Widget
        Public Tag As String
        Public Sub Class_Initialize()
            log = log & "init;"
        End Sub
        Public Sub Class_Terminate()
            log = log & "term;"
        End Sub
    End Class
    Function TestNeverAccessed() As String
        log = ""
        Dim w As New Widget
        Set w = Nothing
        TestNeverAccessed = log
    End Function
    Function TestAccessedFirst() As String
        log = ""
        Dim w As New Widget
        w.Tag = "x"
        Set w = Nothing
        TestAccessedFirst = log
    End Function
    `;
    assert.strictEqual(runFunc(code, 'TestNeverAccessed'), 'init;term;',
        'Dim As New + Set Nothing without any access: Class_Initialize then Class_Terminate must fire');
    console.log('[PASS] Test 15: Dim As New + Set Nothing without prior access fires lifecycle hooks');
    assert.strictEqual(runFunc(code, 'TestAccessedFirst'), 'init;term;',
        'Dim As New + access + Set Nothing: still fires exactly once');
    console.log('[PASS] Test 16: Dim As New + access + Set Nothing fires lifecycle hooks once');
}

console.log('\n✅ Circular Reference & Class_Terminate: 全テスト通過');
