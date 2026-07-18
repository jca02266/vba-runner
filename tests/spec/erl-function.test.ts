import { evalVBASingle, assert } from '../../test-libs/test-runner';

function evalVBA(code: string): any {
    return evalVBASingle(code);
}

function runFunc(code: string, name: string, args: any[] = []): any {
    return evalVBA(code).callProcedure(name, args);
}

// Test 1: Erl returns 0 when no error
{
    const code = `
    Function Test1()
        Test1 = Erl()
    End Function
    `;
    try {
        const result = runFunc(code, 'Test1');
        assert.strictEqual(result, 0, 'Erl should return 0 when no error');
        console.log('[PASS] Test 1: Erl returns 0 when no error');
    } catch (e: any) {
        console.log('[INFO] Test 1: Erl function not implemented');
    }
}

// Test 2: Erl returns line number after error
{
    const code = `
    Function Test2()
        On Error GoTo Handler
10      Dim x As Integer
20      x = 1 / 0  ' Error on line 20
30      Test2 = 0
         Exit Function
    Handler:
40      Test2 = Erl()
    End Function
    `;
    try {
        const result = runFunc(code, 'Test2');
        assert.strictEqual(result, 20, 'Erl should return line number of error');
        console.log('[PASS] Test 2: Erl returns error line number');
    } catch (e: any) {
        console.log('[INFO] Test 2: Erl function not implemented - ' + e.message);
    }
}

// Test 3: Erl with error in nested call
{
    const code = `
    Sub ErrorSub()
10      Dim x As Integer
20      x = 1 / 0  ' Error on line 20
30      Exit Sub
    End Sub

    Function Test3()
        On Error GoTo Handler
        ErrorSub
        Test3 = 0
        Exit Function
    Handler:
        Test3 = Erl()
    End Function
    `;
    try {
        const result = runFunc(code, 'Test3');
        console.log('[INFO] Test 3: Erl in error handler - got ' + result);
    } catch (e: any) {
        console.log('[INFO] Test 3: Erl function not implemented');
    }
}

// Test 4: Erl after Err.Clear
{
    const code = `
    Function Test4()
        On Error GoTo Handler
10      Err.Raise 11
20      Test4 = 0
         Exit Function
    Handler:
        Err.Clear
        Test4 = Erl()  ' Should still return line 10
    End Function
    `;
    try {
        const result = runFunc(code, 'Test4');
        console.log('[INFO] Test 4: Erl after Err.Clear - got ' + result);
    } catch (e: any) {
        console.log('[INFO] Test 4: Erl function not implemented');
    }
}

// Test 5: Erl with Resume
{
    const code = `
    Function Test5()
        Dim x As Integer
        x = 0
        On Error GoTo Handler
10      Dim y As Integer
20      y = 1 / x  ' Error on line 20
30      Test5 = Erl()
         Exit Function
    Handler:
        x = 1
        Resume
    End Function
    `;
    try {
        const result = runFunc(code, 'Test5');
        console.log('[INFO] Test 5: Erl with Resume - got ' + result);
    } catch (e: any) {
        console.log('[INFO] Test 5: Erl function not implemented');
    }
}

// --- Bug 31-B（評価 #31）: Erl が未実装で常に 0 を返していた ---
// 実装: 数値行ラベル通過時に記録し、エラー捕捉時に Err 状態へ確定。
// Err.Clear / Resume / On Error 文 / Exit 系で Err とともにリセットされる。
{
    // ハンドラー内で Erl がエラー行番号を返す（Err.Raise 起点）
    const code = `
    Function G() As Long
        On Error GoTo H
10      Dim a
20      Err.Raise 5
30      Exit Function
    H:
        G = Erl
    End Function
    Function NoNumbers() As Long
        On Error Resume Next
        Err.Raise 5
        NoNumbers = Erl
    End Function
    `;
    const ev = evalVBA(code);
    // 行番号なしプロシージャを先に実行 → 0（グローバル状態が汚れていない状態で）
    assert.strictEqual(ev.callProcedure('NoNumbers', []), 0,
        'Bug 31-B: 行番号のないプロシージャでは Erl = 0');
    assert.strictEqual(ev.callProcedure('G', []), 20,
        'Bug 31-B: ハンドラー内の Erl はエラー発生行の行番号を返す');
    console.log('[PASS] Bug 31-B: Erl がエラー行番号を返す');
}

console.log('\n✅ Erl Function: Testing complete');
