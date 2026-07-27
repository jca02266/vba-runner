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

// `Shared` is a lock clause by itself, not `Lock Shared`.
{
    const vfs4 = new MemoryFileSystem();
    const code = `
        Sub Test()
            Open "shared.dat" For Binary Access Read Write Shared As #1
            Close #1
        End Sub
    `;
    let threw = false;
    try {
        evalVBASingle(code, { fs: vfs4 }).callProcedure('Test', []);
    } catch {
        threw = true;
    }
    assert.strictEqual(threw, false, 'Open ... Access Read Write Shared は Parse エラーにならない');
}
console.log('[PASS] Open Shared lock clause');

// --- Bug 26-7: GetAttr / SetAttr, vbNormal/vbReadOnly/vbDirectory 定数 ---
{
    const attrVfs = new MemoryFileSystem();
    attrVfs.writeFileSync('/sandbox/dummy', '');
    const ev = evalVBASingle('', { fs: attrVfs });
    assert.strictEqual(ev.evalExpression('GetAttr("dummy")'), 0, '通常ファイルは vbNormal');
    assert.strictEqual(ev.evalExpression('vbNormal'), 0, 'vbNormal = 0');
    assert.strictEqual(ev.evalExpression('vbReadOnly'), 1, 'vbReadOnly = 1');
    assert.strictEqual(ev.evalExpression('vbHidden'), 2, 'vbHidden = 2');
    assert.strictEqual(ev.evalExpression('vbSystem'), 4, 'vbSystem = 4');
    assert.strictEqual(ev.evalExpression('vbDirectory'), 16, 'vbDirectory = 16');
    assert.strictEqual(ev.evalExpression('vbArchive'), 32, 'vbArchive = 32');
}
console.log('[PASS] Bug 26-7: GetAttr/SetAttr + ファイル属性定数');

// GetAttr / SetAttr のVFS属性ビットを実体化し、ファイルとディレクトリを区別する。
{
    const attrsVfs = new MemoryFileSystem();
    attrsVfs.writeFileSync('/sandbox/attrs.txt', 'x');
    attrsVfs.mkdirSync('/sandbox/attrs-dir', { recursive: true });
    const attrsEv = evalVBASingle(`
        Public beforeAttrs, afterAttrs, directoryAttrs
        Sub TestAttrs()
            beforeAttrs = GetAttr("attrs.txt")
            SetAttr "attrs.txt", vbReadOnly Or vbHidden
            afterAttrs = GetAttr("attrs.txt")
            directoryAttrs = GetAttr("attrs-dir")
        End Sub
    `, { fs: attrsVfs });
    attrsEv.callProcedure('TestAttrs', []);
    assert.strictEqual(attrsEv.env.get('beforeattrs'), 0, '通常ファイルは vbNormal');
    assert.strictEqual(attrsEv.env.get('afterattrs'), 3, 'SetAttr は ReadOnly と Hidden を保持');
    assert.strictEqual(attrsEv.env.get('directoryattrs'), 16, 'ディレクトリは vbDirectory');
}
console.log('[PASS] GetAttr/SetAttr: VFS属性ビット');

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

// --- Bug BM: Write# が vbaNull を #NULL# として書き込めない ---
{
    const ev = evalVBASingle(`
        Public result As String
        Sub Test()
            Open "C:\\\\test\\\\bm.txt" For Output As #1
            Write #1, Null, Empty
            Close #1
            Open "C:\\\\test\\\\bm.txt" For Input As #1
            Line Input #1, result
            Close #1
        End Sub
    `);
    ev.callProcedure('Test', []);
    assert.strictEqual(ev.env.get('result'), '#NULL#,""', 'Write# Null→#NULL#, Empty→""');
}
console.log('[PASS] Bug BM: Write# Null/Empty encoding');

// Bug 168-A: ANSIテキストI/OはCP932のマルチバイト文字を保持する
{
    const vfs5 = new MemoryFileSystem();
    vfs5.writeFileSync('/sandbox/cp932.txt', Uint8Array.from([0x41, 0x82, 0xa0, 0x0d, 0x0a]) as any);
    const ev = evalVBASingle(`
        Public result As String
        Sub Test()
            Open "cp932.txt" For Input As #1
            Line Input #1, result
            Close #1
        End Sub
    `, { fs: vfs5 });
    ev.callProcedure('Test', []);
    assert.strictEqual(ev.env.get('result'), 'Aあ', 'Line Input はCP932を正しくデコードする');

    const vfs6 = new MemoryFileSystem();
    const writer = evalVBASingle(`
        Sub Test()
            Open "cp932-out.txt" For Output As #1
            Print #1, "Aあ"
            Close #1
        End Sub
    `, { fs: vfs6 });
    writer.callProcedure('Test', []);
    const fd = vfs6.openSync('/sandbox/cp932-out.txt', 'r');
    const bytes = new Uint8Array(5);
    const read = vfs6.readSync(fd, bytes, 0, bytes.length, 0);
    vfs6.closeSync(fd);
    assert.deepStrictEqual(Array.from(bytes.subarray(0, read)), [0x41, 0x82, 0xa0, 0x0d, 0x0a], 'Print はCP932で書き込む');
    console.log('[PASS] Bug 168-A: CP932 text I/O preserves multibyte characters');
}

// Bug 169-A: Input # は Write # の行境界を跨いで値を消費する
{
    const vfs7 = new MemoryFileSystem();
    const ev = evalVBASingle(`
        Public result As String
        Sub Test()
            Open "records.txt" For Output As #1
            Write #1, "Aあ", 101
            Write #1, "B漢", 202
            Close #1
            Dim first As String, second As String, n1 As Long, n2 As Long
            Open "records.txt" For Input As #1
            Input #1, first, n1, second, n2
            result = first & "|" & n1 & "|" & second & "|" & n2 & "|" & CStr(EOF(1))
            Close #1
        End Sub
    `, { fs: vfs7 });
    ev.callProcedure('Test', []);
    assert.strictEqual(ev.env.get('result'), 'Aあ|101|B漢|202|True', 'Input# は複数行のWrite#値を順に読む');
    console.log('[PASS] Bug 169-A: Input# consumes values across line boundaries');
}

console.log('✅ FileSystem: 全テスト通過');
