import { evalVBASingle, assert } from '../../test-libs/test-runner';

console.log("Running Error Handling tests...");

const vbaCode = `
    Sub TestError()
        On Error GoTo ErrorHandler
        Debug.Print "Start"
        Err.Raise 513, "Test", "Custom Error"
        Debug.Print "Back from Error"
        GoTo Done
    ErrorHandler:
        Debug.Print "Caught: " & Err.Number
        Resume Next
    Done:
        Debug.Print "Finished"
    End Sub
    TestError
`;

const lines: string[] = [];
evalVBASingle(vbaCode, { onPrint: (o) => lines.push(o.trim()) });

const expected = ["Start", "Caught: 513", "Back from Error", "Finished"];
expected.forEach((exp, i) => {
    assert.strictEqual(lines[i], exp, `Error Handling line ${i + 1}`);
});

console.log('✅ Error Handling: 全テスト通過');
