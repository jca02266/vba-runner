/**
 * Regression test: Err.Raise inside On Error GoTo handler caused infinite loop
 *
 * Bug: When Err.Raise was called inside a nested block (If/For/While) within an
 * error handler, the inner executeStatements(isTopLevel=false) reset isInErrorHandler=false
 * before re-throwing. The outer executeStatements then saw isInErrorHandler=false and
 * re-entered the same GoTo handler → infinite loop.
 *
 * Fix: Remove isInErrorHandler=false before throw in the "bubble up" path so the outer
 * handler correctly sees isInErrorHandler=true and propagates the error to the caller.
 */
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

// --- Test 1: Err.Raise inside If block within GoTo error handler must propagate to caller ---
{
    const code = `
        Function WithCleanup(throwErr As Boolean)
            Dim sv As Integer
            sv = 42
            On Error GoTo Cleanup
            If throwErr Then Err.Raise 5, "src", "forced error"
            WithCleanup = "ok"
            Exit Function
        Cleanup:
            WithCleanup = "cleaned"
            If Err.Number <> 0 Then
                Err.Raise Err.Number, Err.Source, Err.Description
            End If
        End Function
    `;
    // Normal path: no error
    assert.strictEqual(runFunc(code, 'WithCleanup', [false]), 'ok',
        'GoTo handler: normal path returns "ok"');

    // Error path: Err.Raise in handler's If block must propagate, not loop
    let threw = false;
    let errMsg = '';
    try {
        runFunc(code, 'WithCleanup', [true]);
    } catch (e: any) {
        threw = true;
        errMsg = e.message;
    }
    assert.strictEqual(threw, true, 'Err.Raise inside handler If block propagates to caller');
    assert.strictEqual(errMsg, 'forced error', 'Error message is preserved through re-raise');
    console.log('[PASS] Err.Raise inside handler If block propagates to caller (no infinite loop)');
}

// --- Test 2: SheetsInNewWorkbook pattern (LibBook.bas の実パターン) ---
// Error handler restores state then re-raises — sheetsinnewworkbook must be restored even on error
{
    const code = `
        Function NewWorkbook(SheetCount As Integer)
            Dim sv As Integer
            sv = Application.SheetsInNewWorkbook
            On Error GoTo Cleanup
            Application.SheetsInNewWorkbook = SheetCount
            Set NewWorkbook = Workbooks.Add
            Exit Function
        Cleanup:
            Application.SheetsInNewWorkbook = sv
            If Err.Number <> 0 Then
                Err.Raise Err.Number, Err.Source, Err.Description
            End If
        End Function
    `;
    const mockApp = { sheetsinnewworkbook: 4 };
    const mockWorkbooks = { add: () => { throw new Error('Cannot create workbook'); } };

    const ev = evalVBA(code);
    ev.env.set('application', mockApp);
    ev.env.set('workbooks', mockWorkbooks);

    let threw = false;
    try {
        ev.callProcedure('NewWorkbook', [2]);
    } catch {
        threw = true;
    }

    assert.strictEqual(threw, true, 'Error propagates to caller after cleanup');
    assert.strictEqual(mockApp.sheetsinnewworkbook, 4,
        'SheetsInNewWorkbook restored to original value (4) even on error');
    console.log('[PASS] Cleanup handler restores state and re-raises without looping');
}

// --- Test 3: Err.Raise inside For loop within handler must also propagate ---
{
    const code = `
        Function WithForInHandler()
            On Error GoTo Handler
            Err.Raise 5, "src", "trigger"
            Exit Function
        Handler:
            Dim i As Integer
            For i = 1 To 1
                Err.Raise Err.Number, Err.Source, Err.Description
            Next i
        End Function
    `;
    let threw = false;
    try {
        runFunc(code, 'WithForInHandler', []);
    } catch {
        threw = true;
    }
    assert.strictEqual(threw, true, 'Err.Raise inside For loop in handler propagates to caller');
    console.log('[PASS] Err.Raise inside For loop in handler propagates (no infinite loop)');
}

console.log('\n✅ Error handler re-entry (無限ループ回帰): 全テスト通過');
