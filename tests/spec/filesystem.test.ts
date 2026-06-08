import { evalVBASingle, assert } from '../../test-libs/test-runner';
import { MemoryFileSystem } from '../../src/engine/filesystem';

console.log("Running FileSystem tests...");

const vfs = new MemoryFileSystem();

const vbaCode = `
    Sub TestFile()
        Dim fn As Integer
        fn = FreeFile()
        Open "test.txt" For Output As #fn
        Print #fn, "Hello VBA"
        Print #fn, "Line 2"
        Close #fn

        Dim lineStr As String
        Open "test.txt" For Input As #1
        Line Input #1, lineStr
        Debug.Print lineStr
        Line Input #1, lineStr
        Debug.Print lineStr
        Close #1

        Kill "test.txt"

        Debug.Print Environ("TEMP")
    End Sub
    TestFile
`;

const lines: string[] = [];
evalVBASingle(vbaCode, {
    onPrint: (o) => lines.push(o.trim()),
    fs: vfs,
    env: { temp: "/tmp/vba" },
});

const expected = ["Hello VBA", "Line 2", "/tmp/vba"];
expected.forEach((exp, i) => {
    assert.strictEqual(lines[i], exp, `FileSystem line ${i + 1}`);
});

console.log('✅ FileSystem: 全テスト通過');
