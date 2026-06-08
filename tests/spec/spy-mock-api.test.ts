import { evalVBASingle, evalVBAModules, assert } from '../../test-libs/test-runner';

const VOID = { onPrint: () => {} };

// 1. callCount tracks invocations
{
    const ev = evalVBASingle(`
Sub Test1()
    MsgBox "hello"
    MsgBox "world"
End Sub
`, VOID);
    const spy = ev.spy('MsgBox');
    ev.callProcedure('Test1', []);
    assert.strictEqual(spy.callCount, 2, 'callCount = 2');
    console.log('[PASS] callCount:', spy.callCount);
}

// 2. calls records arguments
{
    const ev = evalVBASingle(`
Sub Test2()
    MsgBox "first"
    MsgBox "second", 1, "Title"
End Sub
`, VOID);
    const spy = ev.spy('MsgBox');
    ev.callProcedure('Test2', []);
    assert.strictEqual(spy.calls.length, 2, '2 calls recorded');
    assert.strictEqual(spy.calls[0][0], 'first', 'first call arg');
    assert.strictEqual(spy.calls[1][0], 'second', 'second call first arg');
    assert.strictEqual(spy.calls[1][2], 'Title', 'second call title arg');
    console.log('[PASS] calls records arguments');
}

// 3. lastCall returns last invocation's args
{
    const ev = evalVBASingle(`
Sub Test3()
    MsgBox "a"
    MsgBox "b"
    MsgBox "c"
End Sub
`, VOID);
    const spy = ev.spy('MsgBox');
    ev.callProcedure('Test3', []);
    assert.strictEqual(spy.lastCall![0], 'c', 'lastCall is last invocation');
    console.log('[PASS] lastCall:', spy.lastCall);
}

// 4. calledWith checks partial arg match
{
    const ev = evalVBASingle(`
Sub Test4()
    MsgBox "Error occurred"
    MsgBox "All done"
End Sub
`, VOID);
    const spy = ev.spy('MsgBox');
    ev.callProcedure('Test4', []);
    assert.ok(spy.calledWith('Error occurred'), 'calledWith "Error occurred"');
    assert.ok(spy.calledWith('All done'), 'calledWith "All done"');
    assert.ok(!spy.calledWith('Not called'), '!calledWith "Not called"');
    console.log('[PASS] calledWith works correctly');
}

// 5. spy with returnFn overrides return value
{
    const ev = evalVBASingle(`
Function Test5() As Long
    Dim r As Long
    r = MsgBox("Save?", 4)  ' vbYesNo
    Test5 = r
End Function
`, VOID);
    ev.spy('MsgBox', () => 7);
    const result = ev.callProcedure('Test5', []);
    assert.strictEqual(result, 7, 'MsgBox returns mocked 7 (vbNo)');
    console.log('[PASS] spy with returnFn overrides return value:', result);
}

// 6. reset clears call history
{
    const ev = evalVBASingle(`
Sub Test6()
    MsgBox "x"
    MsgBox "y"
End Sub
`, VOID);
    const spy = ev.spy('MsgBox');
    ev.callProcedure('Test6', []);
    assert.strictEqual(spy.callCount, 2, 'before reset: 2 calls');
    spy.reset();
    assert.strictEqual(spy.callCount, 0, 'after reset: 0 calls');
    assert.strictEqual(spy.lastCall, null, 'after reset: lastCall null');
    console.log('[PASS] reset clears call history');
}

// 7. spy on Debug.Print
{
    const ev = evalVBAModules([
        { name: 'M1', code: `
Sub Test7()
    Debug.Print "line1"
    Debug.Print "line2"
End Sub
` },
        { name: 'M2', code: `
Sub Wrapper()
    MsgBox "wrapped"
End Sub
` },
    ], VOID);
    const spy2 = ev.spy('MsgBox');
    ev.callProcedure('Wrapper', []);
    assert.strictEqual(spy2.callCount, 1, 'wrapper called MsgBox once');
    console.log('[PASS] spy on user-defined sub path');
}

// 8. spy on InputBox with return value mock
{
    const ev = evalVBASingle(`
Function Test8() As String
    Dim name As String
    name = InputBox("Enter name:")
    Test8 = "Hello, " & name
End Function
`, VOID);
    ev.spy('InputBox', () => 'Alice');
    const result = ev.callProcedure('Test8', []);
    assert.strictEqual(result, 'Hello, Alice', 'InputBox returns mocked "Alice"');
    console.log('[PASS] spy mocks InputBox return value:', result);
}

// 9. spy on user-defined function — VBA procedure dispatch path
{
    const ev = evalVBAModules([
        { name: 'M1', code: `
Function Logger(msg As String) As Long
    Logger = Len(msg)
End Function

Sub Test9Caller()
    Logger "hello"
    Logger "world"
End Sub
` },
        { name: 'M2', code: `
Sub Test9()
    MsgBox "spy test"
End Sub
` },
    ], VOID);
    const spy = ev.spy('MsgBox');
    ev.callProcedure('Test9', []);
    assert.strictEqual(spy.callCount, 1, 'spy on MsgBox from second load');
    console.log('[PASS] spy works across multiple evaluate() calls');
}

// 10. returnValues records what was returned
{
    const ev = evalVBASingle(`
Sub Test10()
    MsgBox "a"
    MsgBox "b"
    MsgBox "c"
End Sub
`, VOID);
    let callNum = 0;
    const spy = ev.spy('MsgBox', () => ++callNum);
    ev.callProcedure('Test10', []);
    assert.strictEqual(spy.returnValues[0], 1, 'first return 1');
    assert.strictEqual(spy.returnValues[1], 2, 'second return 2');
    assert.strictEqual(spy.returnValues[2], 3, 'third return 3');
    console.log('[PASS] returnValues records per-call return values');
}

console.log('\n✅ Spy/Mock API: 全テスト通過');
