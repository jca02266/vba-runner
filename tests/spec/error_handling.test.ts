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

// Err.Raise without Source and implicit runtime errors must not retain the
// custom Source from an earlier error.
const sourceLines: string[] = [];
evalVBASingle(`
    Sub TestSource()
        On Error Resume Next
        Err.Raise 5, "CustomSource", "custom"
        Debug.Print Err.Source
        Err.Raise 6
        Debug.Print Err.Source
        Dim x As Long
        x = 1 / 0
        Debug.Print Err.Source
    End Sub
    TestSource
`, { onPrint: (o) => sourceLines.push(o.trim()) });

assert.deepStrictEqual(sourceLines, ["CustomSource", "VBAProject", "Module1"],
    'Err.Source is refreshed for omitted and implicit runtime errors');
console.log('✅ Err.Source refresh regression passed');
