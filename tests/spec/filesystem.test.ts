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

// 仕様バグ修正: Open ... For Append が書き込み開始位置を常に 0 にしていたため、
// 2回目以降の Append が先頭から上書きしてしまい、追記のたびにファイル内容が
// 消えていた（MemoryFileSystem.openSync）。
{
    const vfs2 = new MemoryFileSystem();
    const appendCode = `
        Sub AppendLine(line As String)
            Dim fn As Integer
            fn = FreeFile()
            Open "append.txt" For Append As #fn
            Print #fn, line
            Close #fn
        End Sub
        AppendLine "line1"
        AppendLine "line2"
        AppendLine "line3"

        Dim fn2 As Integer, s As String, result As String
        fn2 = FreeFile()
        Open "append.txt" For Input As #fn2
        Do While Not EOF(fn2)
            Line Input #fn2, s
            result = result & s & "|"
        Loop
        Close #fn2
        Debug.Print result
    `;
    const appendLines: string[] = [];
    evalVBASingle(appendCode, { onPrint: (o) => appendLines.push(o.trim()), fs: vfs2 });
    assert.strictEqual(appendLines[0], 'line1|line2|line3|', 'For Append は毎回末尾に追記される（先頭上書きしない)');
}
console.log('[PASS] Open For Append は毎回末尾に追記される');

console.log('✅ FileSystem: 全テスト通過');
