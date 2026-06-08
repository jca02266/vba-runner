import { evalVBASingle, assert } from '../../test-libs/test-runner';

function evalVBA(code: string): any {
    return evalVBASingle(code);
}

function runFunc(code: string, name: string, args: any[] = []): any {
    return evalVBA(code).callProcedure(name, args);
}

// --- Do While...Loop (pre-condition while) ---
{
    const code = `
Function DoWhilePre()
    Dim i As Integer
    i = 0
    Do While i < 5
        i = i + 1
    Loop
    DoWhilePre = i
End Function
`;
    assert.strictEqual(runFunc(code, 'DoWhilePre'), 5, 'Do While...Loop: counts to 5');
    console.log('[PASS] Do While...Loop (pre-condition)');
}

// --- Do While...Loop: zero iterations when condition is false initially ---
{
    const code = `
Function DoWhileZero()
    Dim i As Integer
    i = 10
    Do While i < 5
        i = i + 1
    Loop
    DoWhileZero = i
End Function
`;
    assert.strictEqual(runFunc(code, 'DoWhileZero'), 10, 'Do While...Loop: 0 iterations when false initially');
    console.log('[PASS] Do While...Loop: zero iterations');
}

// --- Do Until...Loop (pre-condition until) ---
{
    const code = `
Function DoUntilPre()
    Dim i As Integer
    i = 0
    Do Until i >= 5
        i = i + 1
    Loop
    DoUntilPre = i
End Function
`;
    assert.strictEqual(runFunc(code, 'DoUntilPre'), 5, 'Do Until...Loop: counts to 5');
    console.log('[PASS] Do Until...Loop (pre-condition)');
}

// --- Do Until...Loop: zero iterations when condition is true initially ---
{
    const code = `
Function DoUntilZero()
    Dim i As Integer
    i = 10
    Do Until i >= 5
        i = i + 1
    Loop
    DoUntilZero = i
End Function
`;
    assert.strictEqual(runFunc(code, 'DoUntilZero'), 10, 'Do Until...Loop: 0 iterations when true initially');
    console.log('[PASS] Do Until...Loop: zero iterations');
}

// --- Do...Loop While (post-condition while) ---
{
    const code = `
Function DoLoopWhile()
    Dim i As Integer
    i = 0
    Do
        i = i + 1
    Loop While i < 5
    DoLoopWhile = i
End Function
`;
    assert.strictEqual(runFunc(code, 'DoLoopWhile'), 5, 'Do...Loop While: counts to 5');
    console.log('[PASS] Do...Loop While (post-condition)');
}

// --- Do...Loop While: executes at least once even when false initially ---
{
    const code = `
Function DoLoopWhileOnce()
    Dim i As Integer
    i = 10
    Do
        i = i + 1
    Loop While i < 5
    DoLoopWhileOnce = i
End Function
`;
    assert.strictEqual(runFunc(code, 'DoLoopWhileOnce'), 11, 'Do...Loop While: executes once even when false initially');
    console.log('[PASS] Do...Loop While: executes at least once');
}

// --- Do...Loop Until (post-condition until) ---
{
    const code = `
Function DoLoopUntil()
    Dim i As Integer
    i = 0
    Do
        i = i + 1
    Loop Until i >= 5
    DoLoopUntil = i
End Function
`;
    assert.strictEqual(runFunc(code, 'DoLoopUntil'), 5, 'Do...Loop Until: counts to 5');
    console.log('[PASS] Do...Loop Until (post-condition)');
}

// --- Do...Loop Until: executes at least once even when true initially ---
{
    const code = `
Function DoLoopUntilOnce()
    Dim i As Integer
    i = 10
    Do
        i = i + 1
    Loop Until i >= 5
    DoLoopUntilOnce = i
End Function
`;
    assert.strictEqual(runFunc(code, 'DoLoopUntilOnce'), 11, 'Do...Loop Until: executes once even when true initially');
    console.log('[PASS] Do...Loop Until: executes at least once');
}

// --- Do...Loop (infinite with Exit Do) ---
{
    const code = `
Function DoInfinite()
    Dim i As Integer
    i = 0
    Do
        i = i + 1
        If i >= 5 Then Exit Do
    Loop
    DoInfinite = i
End Function
`;
    assert.strictEqual(runFunc(code, 'DoInfinite'), 5, 'Do...Loop: infinite with Exit Do');
    console.log('[PASS] Do...Loop (infinite with Exit Do)');
}

// --- Exit Do in pre-condition loop ---
{
    const code = `
Function ExitDoWhile()
    Dim i As Integer
    i = 0
    Do While i < 100
        i = i + 1
        If i = 3 Then Exit Do
    Loop
    ExitDoWhile = i
End Function
`;
    assert.strictEqual(runFunc(code, 'ExitDoWhile'), 3, 'Exit Do in Do While...Loop');
    console.log('[PASS] Exit Do in Do While...Loop');
}

// --- Exit Do in post-condition loop ---
{
    const code = `
Function ExitDoLoopUntil()
    Dim i As Integer
    i = 0
    Do
        i = i + 1
        If i = 3 Then Exit Do
    Loop Until i >= 100
    ExitDoLoopUntil = i
End Function
`;
    assert.strictEqual(runFunc(code, 'ExitDoLoopUntil'), 3, 'Exit Do in Do...Loop Until');
    console.log('[PASS] Exit Do in Do...Loop Until');
}

console.log('\n✅ Do Statement: 全テスト通過');
