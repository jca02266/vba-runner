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

// --- Bug 26-3: Open For Random Len = N が Parse エラーにならないこと ---
{
    const vfs3 = new MemoryFileSystem();
    const code = `
        Sub Test()
            Open "rec.dat" For Random As #1 Len = 20
            Close #1
        End Sub
    `;
    let threw = false;
    try {
        evalVBASingle(code, { fs: vfs3 }).callProcedure('Test', []);
    } catch {
        threw = true;
    }
    assert.strictEqual(threw, false, 'Open For Random Len = 20 は Parse エラーにならない');
}
console.log('[PASS] Bug 26-3: Open For Random Len = N');

// --- Bug 26-7: GetAttr / SetAttr スタブ, vbNormal/vbReadOnly/vbDirectory 定数 ---
{
    const ev = evalVBASingle('');
    assert.strictEqual(ev.evalExpression('GetAttr("dummy")'), 0, 'GetAttr は 0 (vbNormal) を返す');
    assert.strictEqual(ev.evalExpression('vbNormal'), 0, 'vbNormal = 0');
    assert.strictEqual(ev.evalExpression('vbReadOnly'), 1, 'vbReadOnly = 1');
    assert.strictEqual(ev.evalExpression('vbHidden'), 2, 'vbHidden = 2');
    assert.strictEqual(ev.evalExpression('vbSystem'), 4, 'vbSystem = 4');
    assert.strictEqual(ev.evalExpression('vbDirectory'), 16, 'vbDirectory = 16');
    assert.strictEqual(ev.evalExpression('vbArchive'), 32, 'vbArchive = 32');
}
console.log('[PASS] Bug 26-7: GetAttr/SetAttr スタブ + ファイル属性定数');

// --- Bug BL: Input # が #TRUE#/#FALSE# を Boolean として読み込めない ---
{
    const ev = evalVBASingle(`
        Public ws, wi, wb1, wb2
        Sub Test()
            Open "C:\\\\test\\\\bl.txt" For Output As #1
            Write #1, "hello", 42, True, False
            Close #1
            Open "C:\\\\test\\\\bl.txt" For Input As #1
            Dim s As String, n As Long, b1 As Boolean, b2 As Boolean
            Input #1, s, n, b1, b2
            ws = s : wi = n : wb1 = b1 : wb2 = b2
            Close #1
        End Sub
    `);
    ev.callProcedure('Test', []);
    assert.strictEqual(ev.env.get('ws'), 'hello', 'Input# string');
    assert.strictEqual(ev.env.get('wi'), 42, 'Input# number');
    assert.strictEqual((ev.env.get('wb1') as any).value, -1, 'Input# #TRUE# → Boolean True');
    assert.strictEqual((ev.env.get('wb2') as any).value, 0, 'Input# #FALSE# → Boolean False');
}
console.log('[PASS] Bug BL: Input# #TRUE#/#FALSE# Boolean パース');

console.log('✅ FileSystem: 全テスト通過');
