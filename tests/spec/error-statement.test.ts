import { evalVBASingle, assert } from '../../test-libs/test-runner';

function runFunc(code: string, name: string, args: any[] = []): any {
    const ev = evalVBASingle(code, { onPrint: () => {} });
    return ev.callProcedure(name, args);
}

// 1. Error Statement
{
    const code = `
        Function TestError()
            On Error GoTo Handler
            Error 11 ' Division by zero
            TestError = "Failed"
            Exit Function
        Handler:
            TestError = "Caught " & Err.Number
        End Function
    `;
    assert.strictEqual(runFunc(code, 'TestError'), "Caught 11", 'Error statement caught by handler');
}

console.log('\n✅ Error Statement: 全テスト通過');
