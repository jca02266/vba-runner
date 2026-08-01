/**
 * CreateObject / New ClassName の組み込みオブジェクトテスト
 *
 * 対象:
 *   - Scripting.Dictionary
 *   - Scripting.FileSystemObject
 *   - MSXML2.XMLHTTP / Microsoft.XMLHTTP
 *   - ADODB.Stream
 *
 * いずれも `CreateObject(progId)` と `New ClassName` の両方で同じファクトリが
 * 使われることを検証する（参照設定相当）。
 *
 * ファイル操作は VFS (MemoryFileSystem) を使用して、ホスト OS に副作用を
 * 与えずに検証する。
 */
import { evalVBASingle, assert, vbaTrue, vbaFalse } from '../../test-libs/test-runner';
import { MemoryFileSystem } from '../../src/engine/filesystem';

console.log('[Test Suite] CreateObject / New 構文の組み込みオブジェクト検証');

// =============================================================================
// 1. Scripting.Dictionary
// =============================================================================
{
    const code = `
        Function TestDictionary()
            Dim dict
            Set dict = CreateObject("Scripting.Dictionary")
            dict.Add "A", 100
            dict.Add "B", 200

            Dim res
            res = 0
            If dict.Exists("A") Then res = res + dict.Item("A")
            If dict.Exists("B") Then res = res + dict.Item("B")

            dict.Remove "A"
            If Not dict.Exists("A") Then res = res + 50

            TestDictionary = res + dict.Count
        End Function

        Function TestNewDictionary()
            Dim d As New Dictionary
            d.Add "A", 10
            d.Add "B", 20
            TestNewDictionary = d.Count
        End Function

        Function TestSetNewDictionary()
            Dim d As Dictionary
            Set d = New Dictionary
            d.Add "key", 42
            TestSetNewDictionary = d.Item("key")
        End Function
    `;
    const ev = evalVBASingle(code);
    assert.strictEqual(ev.callProcedure('TestDictionary', []), 351, 'Dictionary: CreateObject (100+200+50+1)');
    assert.strictEqual(ev.callProcedure('TestNewDictionary', []), 2, 'Dictionary: Dim As New');
    assert.strictEqual(ev.callProcedure('TestSetNewDictionary', []), 42, 'Dictionary: Set = New');
console.log('[PASS] Scripting.Dictionary');

// CompareMode must apply consistently to duplicate detection as well as Exists.
{
    const ev = evalVBASingle(`
        Function CompareModeDuplicate() As String
            Dim dict As Object
            Set dict = CreateObject("Scripting.Dictionary")
            dict.CompareMode = vbTextCompare
            dict.Add "ax-1", 1
            On Error Resume Next
            dict.Add "AX-1", 2
            dict("code") = 1
            dict("CODE") = 2
            CompareModeDuplicate = CStr(dict.Count) & ":" & CStr(Err.Number) & ":" & CStr(dict("code"))
        End Function
    `);
    assert.strictEqual(ev.callProcedure('CompareModeDuplicate', []), '2:457:2',
        'vbTextCompare rejects case-insensitive duplicate keys');
    console.log('[PASS] Dictionary CompareMode duplicate detection');
}
}

// =============================================================================
// 2. Scripting.FileSystemObject
// =============================================================================
{
    const vfs = new MemoryFileSystem();

    const code = `
        Function TestFSOCreateText() As Boolean
            Dim fso
            Set fso = CreateObject("Scripting.FileSystemObject")
            Dim f
            Set f = fso.CreateTextFile("C:\\\\fso1.txt")
            f.WriteLine "Hello"
            f.WriteLine "World"
            f.Close
            TestFSOCreateText = fso.FileExists("C:\\\\fso1.txt")
        End Function

        Function TestFSOReadAll() As String
            Dim fso As New FileSystemObject
            Dim f
            Set f = fso.OpenTextFile("C:\\\\fso1.txt", 1)
            TestFSOReadAll = f.ReadAll
            f.Close
        End Function

        Function TestFSOTextStreamReadSkip() As String
            Dim fso As New FileSystemObject
            Dim f, ts
            Set f = fso.CreateTextFile("C:\\\\stream.txt")
            f.Write "abcdef"
            f.Close
            Set ts = fso.OpenTextFile("C:\\\\stream.txt", 1)
            TestFSOTextStreamReadSkip = ts.Read(2)
            ts.Skip 1
            TestFSOTextStreamReadSkip = TestFSOTextStreamReadSkip & ts.Read(2)
            ts.Close
        End Function

        Function TestFSOTextStreamSkipLine() As String
            Dim fso As New FileSystemObject
            Dim f, ts
            Set f = fso.CreateTextFile("C:\\\\lines.txt")
            f.WriteLine "first"
            f.WriteLine "second"
            f.Close
            Set ts = fso.OpenTextFile("C:\\\\lines.txt", 1)
            ts.SkipLine
            TestFSOTextStreamSkipLine = ts.ReadLine
            ts.Close
        End Function

        Function TestFSOUnicodeWrite() As String
            Dim fso As New FileSystemObject
            Dim f, ts
            Set f = fso.CreateTextFile("C:\\\\unicode.txt", True, True)
            f.Write "猫-日本語"
            f.Close
            Set ts = fso.OpenTextFile("C:\\\\unicode.txt", 1, False, True)
            TestFSOUnicodeWrite = ts.ReadAll
            ts.Close
        End Function

        Function TestFSOUnicodeRead() As String
            Dim fso As New FileSystemObject
            Dim ts
            Set ts = fso.OpenTextFile("C:\\\\preset-unicode.txt", 1, False, True)
            TestFSOUnicodeRead = ts.ReadAll
            ts.Close
        End Function

        Function TestFSOAnsiWrite() As String
            Dim fso As New FileSystemObject
            Dim f, ts
            Set f = fso.CreateTextFile("C:\\\\ansi.txt", True, False)
            f.Write "あ"
            f.Close
            Set ts = fso.OpenTextFile("C:\\\\ansi.txt", 1)
            TestFSOAnsiWrite = ts.ReadAll
            ts.Close
        End Function

        Function TestFSOFolderOps() As Boolean
            Dim fso As FileSystemObject
            Set fso = New FileSystemObject
            fso.CreateFolder "C:\\\\subdir"
            TestFSOFolderOps = fso.FolderExists("C:\\\\subdir")
        End Function

        Function TestFSODeleteFile() As Boolean
            Dim fso As New FileSystemObject
            Dim f
            Set f = fso.CreateTextFile("C:\\\\delme.txt")
            f.Close
            fso.DeleteFile "C:\\\\delme.txt"
            TestFSODeleteFile = Not fso.FileExists("C:\\\\delme.txt")
        End Function

        Function TestFSOCopyMove() As Boolean
            Dim fso As New FileSystemObject
            Dim f
            Set f = fso.CreateTextFile("C:\\\\copy-source.txt")
            f.Write "payload"
            f.Close
            fso.CopyFile "C:\\\\copy-source.txt", "C:\\\\copy-dest.txt"
            If Not fso.FileExists("C:\\\\copy-dest.txt") Then TestFSOCopyMove = False: Exit Function
            fso.MoveFile "C:\\\\copy-dest.txt", "C:\\\\move-dest.txt"
            TestFSOCopyMove = fso.FileExists("C:\\\\move-dest.txt") And Not fso.FileExists("C:\\\\copy-dest.txt")
        End Function

        Function TestFSODeleteMissing() As Long
            Dim fso As New FileSystemObject
            On Error Resume Next
            fso.DeleteFile "C:\\\\missing.txt"
            TestFSODeleteMissing = Err.Number
        End Function

        Function TestFSOPathOps(p As String) As String
            Dim fso As New FileSystemObject
            TestFSOPathOps = fso.GetBaseName(p) & "|" & fso.GetExtensionName(p) & "|" & fso.GetParentFolderName(p)
        End Function

        Function TestFSOReturnedMembers() As String
            Dim fso As New FileSystemObject, stream, file, folder
            Set stream = fso.CreateTextFile("C:\\returned.txt")
            stream.Write "payload"
            stream.Close
            Set file = fso.GetFile("C:\\returned.txt")
            Set folder = fso.GetFolder("C:\\")
            TestFSOReturnedMembers = file.Name & "|" & file.ParentFolder.Name & "|" & folder.Name
            file.Delete
            TestFSOReturnedMembers = TestFSOReturnedMembers & "|" & fso.FileExists("C:\\returned.txt")
        End Function

        Function TestFSOReturnedCollections() As String
            Dim fso As New FileSystemObject, stream, file, folder, child, ts
            fso.CreateFolder "C:\\listing"
            Set stream = fso.CreateTextFile("C:\\listing\\entry.txt")
            stream.Write "entry"
            stream.Close
            fso.CreateFolder "C:\\listing\\child"
            Set file = fso.GetFile("C:\\listing\\entry.txt")
            Set ts = file.OpenAsTextStream(1, -2)
            TestFSOReturnedCollections = ts.ReadAll
            ts.Close
            Set folder = fso.GetFolder("C:\\listing")
            TestFSOReturnedCollections = TestFSOReturnedCollections & "|" & folder.Files.Count & _
                "|" & folder.Files.Item(1).Name & "|" & folder.SubFolders.Count & _
                "|" & folder.SubFolders.Item(1).Name
        End Function
    `;
    const ev = evalVBASingle(code, { fs: vfs });
    assert.strictEqual(ev.callProcedure('TestFSOCreateText', []), vbaTrue, 'FSO: CreateTextFile + FileExists');
    const content = ev.callProcedure('TestFSOReadAll', []) as string;
    assert.strictEqual(content.includes('Hello') && content.includes('World'), true, 'FSO: OpenTextFile + ReadAll');
    assert.strictEqual(ev.callProcedure('TestFSOTextStreamReadSkip', []), 'abde', 'FSO TextStream: Read + Skip');
    assert.strictEqual(ev.callProcedure('TestFSOTextStreamSkipLine', []), 'second', 'FSO TextStream: SkipLine');
    assert.strictEqual(ev.callProcedure('TestFSOUnicodeWrite', []), '猫-日本語', 'FSO TextStream: UTF-16 write/read');
    vfs.writeFileSync('/sandbox/c/preset-unicode.txt', new Uint8Array([
        0xff, 0xfe, 0x2b, 0x00, 0x28, 0x67, 0x0d, 0x00, 0x0a, 0x00,
    ]));
    assert.strictEqual(ev.callProcedure('TestFSOUnicodeRead', []), '+木\r\n', 'FSO TextStream: UTF-16 input');
    assert.strictEqual(ev.callProcedure('TestFSOAnsiWrite', []), 'あ', 'FSO TextStream: CP932 write/read');
    const unicodeFd = vfs.openSync('/sandbox/c/unicode.txt', 'r');
    const unicodeBytes = new Uint8Array(4);
    vfs.readSync(unicodeFd, unicodeBytes, 0, unicodeBytes.length, null);
    vfs.closeSync(unicodeFd);
    assert.deepStrictEqual(Array.from(unicodeBytes), [0xff, 0xfe, 0x2b, 0x73], 'FSO TextStream: UTF-16LE BOM');
    const ansiFd = vfs.openSync('/sandbox/c/ansi.txt', 'r');
    const ansiBytes = new Uint8Array(2);
    vfs.readSync(ansiFd, ansiBytes, 0, ansiBytes.length, null);
    vfs.closeSync(ansiFd);
    assert.deepStrictEqual(Array.from(ansiBytes), [0x82, 0xa0], 'FSO TextStream: CP932 bytes');
    assert.strictEqual(ev.callProcedure('TestFSOFolderOps', []), vbaTrue, 'FSO: CreateFolder + FolderExists');
    assert.strictEqual(ev.callProcedure('TestFSODeleteFile', []), vbaTrue, 'FSO: DeleteFile');
    assert.strictEqual(ev.callProcedure('TestFSOCopyMove', []), vbaTrue, 'FSO: CopyFile + MoveFile');
    assert.strictEqual(ev.callProcedure('TestFSODeleteMissing', []), 53, 'FSO: missing DeleteFile is Error 53');
    assert.strictEqual(
        ev.callProcedure('TestFSOPathOps', ['C:\\path\\to\\report.xlsx']),
        'report|xlsx|C:\\path\\to',
        'FSO: GetBaseName / GetExtensionName / GetParentFolderName'
    );
    assert.strictEqual(
        ev.callProcedure('TestFSOReturnedMembers', []),
        'returned.txt|||False',
        'FSO: GetFile/GetFolder share returned-object members and operations'
    );
    assert.strictEqual(
        ev.callProcedure('TestFSOReturnedCollections', []),
        'entry|1|entry.txt|1|child',
        'FSO: File.OpenAsTextStream and Folder collections share returned objects'
    );
    console.log('[PASS] Scripting.FileSystemObject');
}

// =============================================================================
// 3. MSXML2.XMLHTTP / Microsoft.XMLHTTP
// =============================================================================
{
    const code = `
        Function TestXmlHttpCreateObject() As Long
            Dim http
            Set http = CreateObject("MSXML2.XMLHTTP")
            http.Open "GET", "http://example.com/", False
            http.SetRequestHeader "Accept", "text/plain"
            http.Send
            TestXmlHttpCreateObject = http.ReadyState
        End Function

        Function TestXmlHttpNew() As Long
            Dim http As New XMLHTTP
            http.Open "POST", "http://example.com/api", True
            http.Send "payload"
            TestXmlHttpNew = http.ReadyState
        End Function

        Function TestMicrosoftXmlHttp() As Long
            Dim http
            Set http = CreateObject("Microsoft.XMLHTTP")
            http.Open "GET", "http://example.com/", False
            http.Send
            TestMicrosoftXmlHttp = http.ReadyState
        End Function
    `;
    const ev = evalVBASingle(code);
    assert.strictEqual(ev.callProcedure('TestXmlHttpCreateObject', []), 4, 'XMLHTTP: CreateObject + Open/SetRequestHeader/Send/ReadyState');
    assert.strictEqual(ev.callProcedure('TestXmlHttpNew', []), 4, 'XMLHTTP: New XMLHTTP');
    assert.strictEqual(ev.callProcedure('TestMicrosoftXmlHttp', []), 4, 'XMLHTTP: Microsoft.XMLHTTP も同じ factory');
    console.log('[PASS] MSXML2.XMLHTTP / Microsoft.XMLHTTP');
}

// =============================================================================
// 4. ADODB.Stream
// =============================================================================
{
    const vfs = new MemoryFileSystem();

    const code = `
        Function TestStreamWriteRead() As String
            Dim s
            Set s = CreateObject("ADODB.Stream")
            s.Open
            s.WriteText "Hello, "
            s.WriteText "VBA Stream"
            s.Position = 0
            TestStreamWriteRead = s.ReadText
            s.Close
        End Function

        Function TestStreamSaveLoad() As String
            Dim s1 As New Stream
            s1.Open
            s1.WriteText "persisted content"
            s1.SaveToFile "C:\\\\stream.txt", 2
            s1.Close

            Dim s2 As Stream
            Set s2 = New Stream
            s2.Open
            s2.LoadFromFile "C:\\\\stream.txt"
            TestStreamSaveLoad = s2.ReadText
            s2.Close
        End Function
    `;
    const ev = evalVBASingle(code, { fs: vfs });
    assert.strictEqual(ev.callProcedure('TestStreamWriteRead', []), 'Hello, VBA Stream', 'Stream: WriteText + ReadText');
    assert.strictEqual(ev.callProcedure('TestStreamSaveLoad', []), 'persisted content', 'Stream: SaveToFile + LoadFromFile');
    console.log('[PASS] ADODB.Stream');
}

// =============================================================================
// 5. Auto-Instantiation 仕様（Dim x As New ClassName）
// =============================================================================
{
    const code = `
        Function TestIsNothingFresh() As Boolean
            Dim a As New Dictionary
            TestIsNothingFresh = (a Is Nothing)
        End Function

        Function TestIsNothingPlain() As Boolean
            Dim b As Dictionary
            TestIsNothingPlain = (b Is Nothing)
        End Function

        Function TestReinstantiation() As Long
            Dim a As New Dictionary
            a.Add "x", 1
            Set a = Nothing
            a.Add "y", 2
            TestReinstantiation = a.Count
        End Function

        Function TestIsNothingAfterSetNothing() As Boolean
            Dim a As New Dictionary
            a.Add "x", 1
            Set a = Nothing
            TestIsNothingAfterSetNothing = (a Is Nothing)
        End Function

        Function TestTypeName() As String
            Dim a As New Dictionary
            TestTypeName = TypeName(a)
        End Function
    `;
    const ev = evalVBASingle(code);
    assert.strictEqual(ev.callProcedure('TestIsNothingFresh', []), vbaFalse, 'Auto-instance: 直後の Is Nothing は False');
    assert.strictEqual(ev.callProcedure('TestIsNothingPlain', []), vbaTrue, '通常 Dim: Is Nothing は True');
    assert.strictEqual(ev.callProcedure('TestReinstantiation', []), 1, 'Auto-instance: Set Nothing 後の再インスタンス化');
    assert.strictEqual(ev.callProcedure('TestIsNothingAfterSetNothing', []), vbaFalse, 'Auto-instance: Set Nothing 後も Is Nothing は False');
    assert.strictEqual(ev.callProcedure('TestTypeName', []), 'Dictionary', 'Auto-instance: TypeName は class 名');
    console.log('[PASS] Auto-Instantiation (Dim As New)');
}

console.log('\n✅ CreateObject: 全テスト通過');
