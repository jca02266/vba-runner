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

        Function TestFSOPathOps(p As String) As String
            Dim fso As New FileSystemObject
            TestFSOPathOps = fso.GetBaseName(p) & "|" & fso.GetExtensionName(p) & "|" & fso.GetParentFolderName(p)
        End Function
    `;
    const ev = evalVBASingle(code, { fs: vfs });
    assert.strictEqual(ev.callProcedure('TestFSOCreateText', []), vbaTrue, 'FSO: CreateTextFile + FileExists');
    const content = ev.callProcedure('TestFSOReadAll', []) as string;
    assert.strictEqual(content.includes('Hello') && content.includes('World'), true, 'FSO: OpenTextFile + ReadAll');
    assert.strictEqual(ev.callProcedure('TestFSOFolderOps', []), vbaTrue, 'FSO: CreateFolder + FolderExists');
    assert.strictEqual(ev.callProcedure('TestFSODeleteFile', []), vbaTrue, 'FSO: DeleteFile');
    assert.strictEqual(
        ev.callProcedure('TestFSOPathOps', ['C:\\path\\to\\report.xlsx']),
        'report|xlsx|C:\\path\\to',
        'FSO: GetBaseName / GetExtensionName / GetParentFolderName'
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
