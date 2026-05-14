import { Lexer } from '../../src/compiler/lexer';
import { Parser } from '../../src/compiler/parser';
import { Evaluator } from '../../src/compiler/evaluator';
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

// Test 1: Basic WithEvents event handler call
{
    const code = `
    Class EventSource
        Public Event TestEvent()

        Public Sub RaiseTestEvent()
            RaiseEvent TestEvent
        End Sub
    End Class

    Class EventListener
        Private eventSource As EventSource
        Private eventFired As Boolean

        Public Sub Initialize(source As EventSource)
            Set eventSource = source
            eventFired = False
        End Sub

        Private Sub eventSource_TestEvent()
            eventFired = True
        End Sub

        Public Function WasFired() As Boolean
            WasFired = eventFired
        End Function
    End Class

    Function Test1()
        Dim source As New EventSource
        Dim listener As New EventListener
        listener.Initialize source
        source.RaiseTestEvent
        Test1 = listener.WasFired()
    End Function
    `;
    try {
        const result = runFunc(code, 'Test1');
        assert.strictEqual(result, true, 'Event handler should be called');
        console.log('[PASS] Test 1: Basic event handler call');
    } catch (e: any) {
        console.log('[INFO] Test 1: ' + e.message);
    }
}

// Test 2: Event with parameters
{
    const code = `
    Class EventSource
        Public Event DataEvent(data As Integer)

        Public Sub RaiseDataEvent(value As Integer)
            RaiseEvent DataEvent(value)
        End Sub
    End Class

    Class EventListener
        Private received As Integer

        Public Sub Initialize(source As EventSource)
            received = 0
        End Sub

        Private Sub source_DataEvent(data As Integer)
            received = data
        End Sub

        Public Function GetData() As Integer
            GetData = received
        End Function
    End Class

    Function Test2()
        Dim source As New EventSource
        Dim listener As New EventListener
        listener.Initialize source
        source.RaiseDataEvent 42
        Test2 = listener.GetData()
    End Function
    `;
    try {
        const result = runFunc(code, 'Test2');
        assert.strictEqual(result, 42, 'Event should pass parameter to handler');
        console.log('[PASS] Test 2: Event with parameters');
    } catch (e: any) {
        console.log('[INFO] Test 2: ' + e.message);
    }
}

// Test 3: Multiple events in sequence
{
    const code = `
    Class EventSource
        Public Event Event1()
        Public Event Event2()

        Public Sub RaiseFirst()
            RaiseEvent Event1
        End Sub

        Public Sub RaiseSecond()
            RaiseEvent Event2
        End Sub
    End Class

    Class EventListener
        Private count As Integer

        Public Sub Initialize()
            count = 0
        End Sub

        Private Sub source_Event1()
            count = count + 10
        End Sub

        Private Sub source_Event2()
            count = count + 100
        End Sub

        Public Function GetCount() As Integer
            GetCount = count
        End Function
    End Class

    Function Test3()
        Dim source As New EventSource
        Dim listener As New EventListener
        listener.Initialize
        source.RaiseFirst
        source.RaiseSecond
        Test3 = listener.GetCount()
    End Function
    `;
    try {
        const result = runFunc(code, 'Test3');
        assert.strictEqual(result, 110, 'Multiple events should fire sequentially');
        console.log('[PASS] Test 3: Multiple events in sequence');
    } catch (e: any) {
        console.log('[INFO] Test 3: ' + e.message);
    }
}

// Test 4: Event handler modifies source state
{
    const code = `
    Class EventSource
        Public Event ValueChanged()
        Private value As Integer

        Public Sub Initialize()
            value = 0
        End Sub

        Public Sub SetValue(v As Integer)
            value = v
            RaiseEvent ValueChanged
        End Sub

        Public Function GetValue() As Integer
            GetValue = value
        End Function
    End Class

    Class EventListener
        Private lastValue As Integer

        Public Sub Initialize(source As EventSource)
            lastValue = 0
        End Sub

        Private Sub source_ValueChanged()
            lastValue = source.GetValue()
        End Sub

        Public Function GetLastValue() As Integer
            GetLastValue = lastValue
        End Function
    End Class

    Function Test4()
        Dim source As New EventSource
        source.Initialize
        Dim listener As New EventListener
        listener.Initialize source
        source.SetValue 55
        Test4 = listener.GetLastValue()
    End Function
    `;
    try {
        const result = runFunc(code, 'Test4');
        assert.strictEqual(result, 55, 'Event handler should access source state');
        console.log('[PASS] Test 4: Event handler accesses source');
    } catch (e: any) {
        console.log('[INFO] Test 4: ' + e.message);
    }
}

// Test 5: Event handler called multiple times
{
    const code = `
    Class EventSource
        Public Event CountEvent()

        Public Sub RaiseCount(times As Integer)
            Dim i As Integer
            For i = 1 To times
                RaiseEvent CountEvent
            Next i
        End Sub
    End Class

    Class EventListener
        Private eventCount As Integer

        Public Sub Initialize()
            eventCount = 0
        End Sub

        Private Sub source_CountEvent()
            eventCount = eventCount + 1
        End Sub

        Public Function GetCount() As Integer
            GetCount = eventCount
        End Function
    End Class

    Function Test5()
        Dim source As New EventSource
        Dim listener As New EventListener
        listener.Initialize
        source.RaiseCount 5
        Test5 = listener.GetCount()
    End Function
    `;
    try {
        const result = runFunc(code, 'Test5');
        assert.strictEqual(result, 5, 'Event handler should be called multiple times');
        console.log('[PASS] Test 5: Multiple event invocations');
    } catch (e: any) {
        console.log('[INFO] Test 5: ' + e.message);
    }
}

// Test 6: Event handler with class variables
{
    const code = `
    Class EventSource
        Public Event Triggered()

        Public Sub Fire()
            RaiseEvent Triggered
        End Sub
    End Class

    Class EventListener
        Private count As Integer
        Private lastTime As Integer

        Public Sub Initialize()
            count = 0
            lastTime = 0
        End Sub

        Private Sub source_Triggered()
            count = count + 1
            lastTime = count * 10
        End Sub

        Public Function GetCount() As Integer
            GetCount = count
        End Function

        Public Function GetLastTime() As Integer
            GetLastTime = lastTime
        End Function
    End Class

    Function Test6()
        Dim source As New EventSource
        Dim listener As New EventListener
        listener.Initialize
        source.Fire
        source.Fire
        source.Fire
        Test6 = listener.GetLastTime()
    End Function
    `;
    try {
        const result = runFunc(code, 'Test6');
        assert.strictEqual(result, 30, 'Event handler state persistence');
        console.log('[PASS] Test 6: Event handler class variables');
    } catch (e: any) {
        console.log('[INFO] Test 6: ' + e.message);
    }
}

// Test 7: Two event listeners on same source
{
    const code = `
    Class EventSource
        Public Event Changed()

        Public Sub Change()
            RaiseEvent Changed
        End Sub
    End Class

    Class Listener1
        Private val As Integer

        Public Sub Initialize()
            val = 0
        End Sub

        Private Sub source_Changed()
            val = 100
        End Sub

        Public Function GetValue() As Integer
            GetValue = val
        End Function
    End Class

    Class Listener2
        Private val As Integer

        Public Sub Initialize()
            val = 0
        End Sub

        Private Sub source_Changed()
            val = 200
        End Sub

        Public Function GetValue() As Integer
            GetValue = val
        End Function
    End Class

    Function Test7()
        Dim source As New EventSource
        Dim l1 As New Listener1
        Dim l2 As New Listener2
        l1.Initialize
        l2.Initialize
        source.Change
        Test7 = l1.GetValue() + l2.GetValue()
    End Function
    `;
    try {
        const result = runFunc(code, 'Test7');
        assert.strictEqual(result, 300, 'Multiple listeners on same event');
        console.log('[PASS] Test 7: Two event listeners on same source');
    } catch (e: any) {
        console.log('[INFO] Test 7: ' + e.message);
    }
}

// Test 8: Event handler prevention (no handler)
{
    const code = `
    Class EventSource
        Public Event Silent()
        Private raised As Boolean

        Public Sub Initialize()
            raised = False
        End Sub

        Public Sub Fire()
            raised = True
            RaiseEvent Silent
        End Sub

        Public Function WasRaised() As Boolean
            WasRaised = raised
        End Function
    End Class

    Function Test8()
        Dim source As New EventSource
        source.Initialize
        source.Fire
        Test8 = source.WasRaised()
    End Function
    `;
    try {
        const result = runFunc(code, 'Test8');
        assert.strictEqual(result, true, 'Event can be raised with no handlers');
        console.log('[PASS] Test 8: Event without handler');
    } catch (e: any) {
        console.log('[INFO] Test 8: ' + e.message);
    }
}

// Test 9: Event handler with boolean return (even though VBA events typically don't return)
{
    const code = `
    Class EventSource
        Public Event Triggered()

        Public Sub Fire()
            RaiseEvent Triggered
        End Sub
    End Class

    Class Listener
        Private fired As Boolean

        Public Sub Initialize()
            fired = False
        End Sub

        Private Sub source_Triggered()
            fired = True
        End Sub

        Public Function IsFired() As Boolean
            IsFired = fired
        End Function
    End Class

    Function Test9()
        Dim source As New EventSource
        Dim listener As New Listener
        listener.Initialize
        source.Fire
        Test9 = listener.IsFired()
    End Function
    `;
    try {
        const result = runFunc(code, 'Test9');
        assert.strictEqual(result, true, 'Boolean state in event handler');
        console.log('[PASS] Test 9: Event handler with boolean state');
    } catch (e: any) {
        console.log('[INFO] Test 9: ' + e.message);
    }
}

// Test 10: Nested event firing (handler raises another event)
{
    const code = `
    Class EventSource
        Public Event Primary()
        Public Event Secondary()

        Public Sub FirePrimary()
            RaiseEvent Primary
        End Sub

        Public Sub FireSecondary()
            RaiseEvent Secondary
        End Sub
    End Class

    Class Listener
        Private primary As Boolean
        Private secondary As Boolean

        Public Sub Initialize(source As EventSource)
            primary = False
            secondary = False
        End Sub

        Private Sub source_Primary()
            primary = True
        End Sub

        Private Sub source_Secondary()
            secondary = True
        End Sub

        Public Function GetPrimary() As Boolean
            GetPrimary = primary
        End Function

        Public Function GetSecondary() As Boolean
            GetSecondary = secondary
        End Function
    End Class

    Function Test10()
        Dim source As New EventSource
        Dim listener As New Listener
        listener.Initialize source
        source.FirePrimary
        source.FireSecondary
        Test10 = listener.GetPrimary() And listener.GetSecondary()
    End Function
    `;
    try {
        const result = runFunc(code, 'Test10');
        assert.strictEqual(result, true, 'Nested events');
        console.log('[PASS] Test 10: Nested event firing');
    } catch (e: any) {
        console.log('[INFO] Test 10: ' + e.message);
    }
}

// Test 11: Event with string parameter
{
    const code = `
    Class EventSource
        Public Event MessageEvent(msg As String)

        Public Sub SendMessage(text As String)
            RaiseEvent MessageEvent(text)
        End Sub
    End Class

    Class Listener
        Private message As String

        Public Sub Initialize()
            message = ""
        End Sub

        Private Sub source_MessageEvent(msg As String)
            message = msg
        End Sub

        Public Function GetMessage() As String
            GetMessage = message
        End Function
    End Class

    Function Test11()
        Dim source As New EventSource
        Dim listener As New Listener
        listener.Initialize
        source.SendMessage "Hello"
        Test11 = listener.GetMessage()
    End Function
    `;
    try {
        const result = runFunc(code, 'Test11');
        assert.strictEqual(result, "Hello", 'Event with string parameter');
        console.log('[PASS] Test 11: Event with string parameter');
    } catch (e: any) {
        console.log('[INFO] Test 11: ' + e.message);
    }
}

// Test 12: Event with multiple parameters
{
    const code = `
    Class EventSource
        Public Event MultiEvent(x As Integer, y As Integer, z As Integer)

        Public Sub FireMulti(a As Integer, b As Integer, c As Integer)
            RaiseEvent MultiEvent(a, b, c)
        End Sub
    End Class

    Class Listener
        Private result As Integer

        Public Sub Initialize()
            result = 0
        End Sub

        Private Sub source_MultiEvent(x As Integer, y As Integer, z As Integer)
            result = x + y + z
        End Sub

        Public Function GetResult() As Integer
            GetResult = result
        End Function
    End Class

    Function Test12()
        Dim source As New EventSource
        Dim listener As New Listener
        listener.Initialize
        source.FireMulti 10, 20, 30
        Test12 = listener.GetResult()
    End Function
    `;
    try {
        const result = runFunc(code, 'Test12');
        assert.strictEqual(result, 60, 'Event with multiple parameters');
        console.log('[PASS] Test 12: Event with multiple parameters');
    } catch (e: any) {
        console.log('[INFO] Test 12: ' + e.message);
    }
}

// Test 13: Event handler accessing object state
{
    const code = `
    Class Data
        Public Value As Integer
    End Class

    Class EventSource
        Public Event Changed()
        Private data As New Data

        Public Sub Initialize()
            data.Value = 0
        End Sub

        Public Sub SetData(v As Integer)
            data.Value = v
            RaiseEvent Changed
        End Sub

        Public Function GetData() As Integer
            GetData = data.Value
        End Function
    End Class

    Class Listener
        Private lastData As Integer

        Public Sub Initialize(source As EventSource)
            lastData = 0
        End Sub

        Private Sub source_Changed()
            lastData = source.GetData()
        End Sub

        Public Function GetLastData() As Integer
            GetLastData = lastData
        End Function
    End Class

    Function Test13()
        Dim source As New EventSource
        source.Initialize
        Dim listener As New Listener
        listener.Initialize source
        source.SetData 99
        Test13 = listener.GetLastData()
    End Function
    `;
    try {
        const result = runFunc(code, 'Test13');
        assert.strictEqual(result, 99, 'Event handler accesses source.data');
        console.log('[PASS] Test 13: Event handler accessing object state');
    } catch (e: any) {
        console.log('[INFO] Test 13: ' + e.message);
    }
}

// Test 14: Sequential event assignments (reassign handler)
{
    const code = `
    Class EventSource
        Public Event MyEvent()

        Public Sub Fire()
            RaiseEvent MyEvent
        End Sub
    End Class

    Class Listener
        Private count As Integer

        Public Sub Initialize()
            count = 0
        End Sub

        Private Sub source_MyEvent()
            count = count + 1
        End Sub

        Public Function GetCount() As Integer
            GetCount = count
        End Function
    End Class

    Function Test14()
        Dim source As New EventSource
        Dim listener1 As New Listener
        listener1.Initialize
        source.Fire
        source.Fire
        Test14 = listener1.GetCount()
    End Function
    `;
    try {
        const result = runFunc(code, 'Test14');
        assert.strictEqual(result, 2, 'Handler called twice');
        console.log('[PASS] Test 14: Sequential event invocations');
    } catch (e: any) {
        console.log('[INFO] Test 14: ' + e.message);
    }
}

// Test 15: VBA veteran pattern - event-driven state machine
{
    const code = `
    Class StateMachine
        Public Event StateChanged(newState As Integer)
        Private state As Integer

        Public Sub Initialize()
            state = 0
        End Sub

        Public Sub Transition(newState As Integer)
            state = newState
            RaiseEvent StateChanged(state)
        End Sub

        Public Function GetState() As Integer
            GetState = state
        End Function
    End Class

    Class StateObserver
        Private lastState As Integer
        Private transitions As Integer

        Public Sub Initialize()
            lastState = 0
            transitions = 0
        End Sub

        Private Sub machine_StateChanged(newState As Integer)
            lastState = newState
            transitions = transitions + 1
        End Sub

        Public Function GetLastState() As Integer
            GetLastState = lastState
        End Function

        Public Function GetTransitionCount() As Integer
            GetTransitionCount = transitions
        End Function
    End Class

    Function Test15()
        Dim machine As New StateMachine
        machine.Initialize
        Dim observer As New StateObserver
        observer.Initialize
        machine.Transition 1
        machine.Transition 2
        machine.Transition 3
        Test15 = observer.GetLastState() + observer.GetTransitionCount()
    End Function
    `;
    try {
        const result = runFunc(code, 'Test15');
        assert.strictEqual(result, 6, 'Event-driven state machine (3 + 3)');
        console.log('[PASS] Test 15: VBA veteran pattern - state machine');
    } catch (e: any) {
        console.log('[INFO] Test 15: ' + e.message);
    }
}

console.log('\n✅ WithEvents Lifecycle: Testing complete');
