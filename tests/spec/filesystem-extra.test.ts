/**
 * FileSystem Functions (Extra) (§6.1.2.5) のテスト
 */
import { evalVBASingle, assert, vbaTrue, vbaFalse } from '../../test-libs/test-runner';
import { MemoryFileSystem } from '../../src/engine/filesystem';
import { NodeFileSystem } from '../../src/engine/node_filesystem';
import * as nodeFs from 'node:fs';
import * as nodePath from 'node:path';
import * as nodeOs from 'node:os';

// Use VFS (MemoryFileSystem) for tests
const vfs = new MemoryFileSystem();

function evalVBA(code: string): any {
    const ev = evalVBASingle(code, { fs: vfs });
    return ev;
}

// すべてのテストを1つのコード内で実行（複数のプロシージャ）
const allCode = String.raw`
    Public s, flen, fdate, posBefore, posAfter, attrFd, attrNum, attrErr, dirNormal, dirDirectory, copyErr, folderCopyResult, folderMoveResult, fileCopyResult, fileMoveResult
    Public eofNullErr, lofNullErr, locNullErr, seekNullErr, openNullErr, mkdirErr, mkdirMissingErr
    Public putRecordNullErr, getRecordNullErr, seekPositionNullErr, rmdirNonEmptyErr, fsoModeErr, killDirectoryErr, deleteFolderErr, openMissingParentErr, copyFolderResult, moveFolderResult, setattrDirectoryErr, deleteFileNoForceErr, deleteFileNoForceExists, deleteFileForceErr, deleteFileForceExists, deleteFileDirectoryErr, fsoFolderExistingErr, fsoFolderMissingParentErr

    Sub Test1()
        ' --- 1. Put / Get (Binary mode) ---
        Open "test_bin.dat" For Binary As #1
        Put #1, , "VBA-Runner"
        Close #1

        Open "test_bin.dat" For Binary As #2
        Get #2, , s
        Close #2

        Debug.Print s
    End Sub

    Sub Test2()
        ' --- 2. FileLen, FileDateTime ---
        flen = FileLen("test_bin.dat")
        fdate = FileDateTime("test_bin.dat")
    End Sub

    Sub Test3()
        ' --- 3. Seek, Kill ---
        Open "test_bin.dat" For Binary As #1
        posBefore = Seek(1)
        Seek #1, 5
        posAfter = Seek(1)
        Close #1
        Kill "test_bin.dat"
    End Sub

    Sub Test4()
        Open "test_attr.dat" For Binary As #1
        attrFd = FileAttr(1, 1)
        attrNum = FileAttr(1, 2)
        Close #1
        Kill "test_attr.dat"
    End Sub

    Sub Test4Invalid()
        Open "test_attr.dat" For Binary As #1
        On Error Resume Next
        attrErr = 0
        attrFd = FileAttr(1, 0)
        attrErr = Err.Number
        Close #1
        Kill "test_attr.dat"
    End Sub

    Sub Test5DirAttributes()
        MkDir "test_dir"
        Open "test_file.txt" For Output As #1
        Print #1, "x"
        Close #1
        dirNormal = Dir("test_dir", vbNormal)
        dirDirectory = Dir("test_dir", vbDirectory)
        Kill "test_file.txt"
        RmDir "test_dir"
    End Sub

    Sub Test6FileCopyExistingDestination()
        Open "copy_src.txt" For Output As #1
        Print #1, "source"
        Close #1
        Open "copy_dst.txt" For Output As #1
        Print #1, "destination"
        Close #1
        On Error Resume Next
        FileCopy "copy_src.txt", "copy_dst.txt"
        copyErr = Err.Number
        On Error GoTo 0
        Kill "copy_src.txt"
        Kill "copy_dst.txt"
    End Sub

    Sub Test7FileNumberNull()
        On Error Resume Next
        Err.Clear
        eofNullErr = Err.Number
        eofNullErr = 0
        eofNullErr = EOF(Null)
        eofNullErr = Err.Number
        Err.Clear
        lofNullErr = LOF(Null)
        lofNullErr = Err.Number
        Err.Clear
        locNullErr = Loc(Null)
        locNullErr = Err.Number
        Err.Clear
        seekNullErr = Seek(Null)
        seekNullErr = Err.Number
        Err.Clear
        Open "null-file.txt" For Output As #Null
        openNullErr = Err.Number
        On Error GoTo 0
    End Sub

    Sub Test8MkDirExistingPath()
        MkDir "existing_dir"
        On Error Resume Next
        MkDir "existing_dir"
        mkdirErr = Err.Number
        On Error GoTo 0
        RmDir "existing_dir"
        On Error Resume Next
        MkDir "missing_parent/child"
        mkdirMissingErr = Err.Number
        On Error GoTo 0
    End Sub

    Sub Test9RecordPositionNull()
        Dim b As Byte
        Open "record-null.bin" For Binary As #1
        On Error Resume Next
        Err.Clear
        Put #1, Null, b
        putRecordNullErr = Err.Number
        Err.Clear
        Get #1, Null, b
        getRecordNullErr = Err.Number
        Err.Clear
        Seek #1, Null
        seekPositionNullErr = Err.Number
        Close #1
        On Error GoTo 0
    End Sub

    Sub Test10RmDirNonEmpty()
        MkDir "nonempty_dir"
        Open "nonempty_dir\child.txt" For Output As #1
        Print #1, "x"
        Close #1
        On Error Resume Next
        Err.Clear
        RmDir "nonempty_dir"
        rmdirNonEmptyErr = Err.Number
        On Error GoTo 0
        Kill "nonempty_dir\child.txt"
        RmDir "nonempty_dir"
    End Sub

    Sub Test11FsoIOModeBoundary()
        Dim fso As Object, stream As Object
        Set fso = CreateObject("Scripting.FileSystemObject")
        On Error Resume Next
        Err.Clear
        Set stream = fso.OpenTextFile("mode-boundary.txt", 3, True, -2)
        fsoModeErr = Err.Number
        On Error GoTo 0
    End Sub

    Sub Test12KillDirectory()
        MkDir "kill_dir"
        On Error Resume Next
        Err.Clear
        Kill "kill_dir"
        killDirectoryErr = Err.Number
        On Error GoTo 0
        RmDir "kill_dir"
    End Sub

    Sub Test13FsoDeleteFolderForce()
        Dim fso As Object, stream As Object
        Set fso = CreateObject("Scripting.FileSystemObject")
        fso.CreateFolder "delete_folder"
        Set stream = fso.CreateTextFile("delete_folder\child.txt")
        stream.Write "x"
        stream.Close
        On Error Resume Next
        Err.Clear
        fso.DeleteFolder "delete_folder"
        deleteFolderErr = Err.Number
        On Error GoTo 0
        fso.DeleteFolder "delete_folder", True
    End Sub

    Sub Test14OpenOutputMissingParent()
        On Error Resume Next
        Err.Clear
        Open "missing_parent\output.txt" For Output As #1
        openMissingParentErr = Err.Number
        If Err.Number = 0 Then Close #1
        On Error GoTo 0
    End Sub

    Sub Test15FsoCopyFolder()
        Dim fso As Object, stream As Object
        Set fso = CreateObject("Scripting.FileSystemObject")
        fso.CreateFolder "copy_folder_source"
        Set stream = fso.CreateTextFile("copy_folder_source\child.txt")
        stream.Write "x"
        stream.Close
        fso.CopyFolder "copy_folder_source", "copy_folder_dest", False
        copyFolderResult = IIf(fso.FolderExists("copy_folder_dest") And fso.FileExists("copy_folder_dest\child.txt"), "ok", "bad")
        fso.DeleteFolder "copy_folder_dest", True
        fso.DeleteFolder "copy_folder_source", True
    End Sub

    Sub Test16FsoMoveFolder()
        Dim fso As Object, stream As Object
        Set fso = CreateObject("Scripting.FileSystemObject")
        fso.CreateFolder "move_folder_source"
        Set stream = fso.CreateTextFile("move_folder_source\child.txt")
        stream.Write "x"
        stream.Close
        fso.MoveFolder "move_folder_source", "move_folder_dest"
        moveFolderResult = IIf((Not fso.FolderExists("move_folder_source")) And fso.FileExists("move_folder_dest\child.txt"), "ok", "bad")
        fso.DeleteFolder "move_folder_dest", True
    End Sub

    Sub Test22FsoPathObjectCopyMove()
        Dim fso As Object, source As Object, child As Object, copied As Object, moved As Object
        Set fso = CreateObject("Scripting.FileSystemObject")
        fso.CreateFolder "path_object_source"
        Set child = fso.CreateTextFile("path_object_source\child.txt")
        child.Write "x"
        child.Close
        Set source = fso.GetFolder("path_object_source")
        source.Copy "path_object_copy"
        folderCopyResult = IIf(fso.FolderExists("path_object_copy") And fso.FileExists("path_object_copy\child.txt"), "ok", "bad")
        Set copied = fso.GetFile("path_object_copy\child.txt")
        copied.Copy "path_object_file_copy.txt"
        fileCopyResult = IIf(fso.FileExists("path_object_file_copy.txt"), "ok", "bad")
        Set moved = fso.GetFolder("path_object_copy")
        moved.Move "path_object_move"
        folderMoveResult = IIf((Not fso.FolderExists("path_object_copy")) And fso.FileExists("path_object_move\child.txt"), "ok", "bad")
        Set copied = fso.GetFile("path_object_file_copy.txt")
        copied.Move "path_object_file_move.txt"
        fileMoveResult = IIf((Not fso.FileExists("path_object_file_copy.txt")) And fso.FileExists("path_object_file_move.txt"), "ok", "bad")
        fso.DeleteFolder "path_object_source", True
        fso.DeleteFolder "path_object_move", True
        fso.DeleteFile "path_object_file_move.txt", True
    End Sub

    Sub Test17SetAttrDirectory()
        Open "attr-file.txt" For Output As #1
        Close #1
        On Error Resume Next
        Err.Clear
        SetAttr "attr-file.txt", vbDirectory
        setattrDirectoryErr = Err.Number
        On Error GoTo 0
        Kill "attr-file.txt"
    End Sub

    Sub Test18FsoCreateFolderBoundaries()
        Dim fso As Object
        Set fso = CreateObject("Scripting.FileSystemObject")
        fso.CreateFolder "fso-existing"
        On Error Resume Next
        Err.Clear
        fso.CreateFolder "fso-existing"
        fsoFolderExistingErr = Err.Number
        Err.Clear
        fso.CreateFolder "fso-missing-parent\child"
        fsoFolderMissingParentErr = Err.Number
        On Error GoTo 0
        fso.DeleteFolder "fso-existing", True
    End Sub

    Sub Test20FsoDeleteFileForce()
        Dim fso As Object, stream As Object
        Set fso = CreateObject("Scripting.FileSystemObject")
        Set stream = fso.CreateTextFile("readonly-file.txt")
        stream.Close
        SetAttr "readonly-file.txt", vbReadOnly
        On Error Resume Next
        Err.Clear
        fso.DeleteFile "readonly-file.txt", False
        deleteFileNoForceErr = Err.Number
        deleteFileNoForceExists = fso.FileExists("readonly-file.txt")
        Err.Clear
        fso.DeleteFile "readonly-file.txt", True
        deleteFileForceErr = Err.Number
        deleteFileForceExists = fso.FileExists("readonly-file.txt")
        On Error GoTo 0
    End Sub

    Sub Test21FsoDeleteFileDirectory()
        Dim fso As Object
        Set fso = CreateObject("Scripting.FileSystemObject")
        fso.CreateFolder "delete-file-dir"
        On Error Resume Next
        Err.Clear
        fso.DeleteFile "delete-file-dir", True
        deleteFileDirectoryErr = Err.Number
        On Error GoTo 0
        fso.DeleteFolder "delete-file-dir", True
    End Sub
`;

const ev = evalVBA(allCode);

// Test 1: Put/Get
ev.callProcedure('Test1', []);
assert.strictEqual(ev.env.get('s'), "VBA-Runner", 'Put/Get (Binary)');
console.log('[PASS] Put/Get (Binary)');

// Test 2: FileLen, FileDateTime
ev.callProcedure('Test2', []);
assert.ok(ev.env.get('flen') > 0, 'FileLen > 0');
assert.ok(ev.env.get('fdate') && ev.env.get('fdate').__isVbaDate__, 'FileDateTime');
console.log('[PASS] FileLen, FileDateTime');

// Test 3: Seek, Kill
ev.callProcedure('Test3', []);
assert.strictEqual(ev.env.get('posbefore'), 1, 'Seek(1) 初期位置');
assert.strictEqual(ev.env.get('posafter'), 5, 'Seek #1, 5');
// VFS 内のファイル存在確認
assert.strictEqual(vfs.existsSync('/test_bin.dat'), false, 'Kill');
console.log('[PASS] Seek, Kill');

// Test 4: FileAttr return type and validation
ev.callProcedure('Test4', []);
assert.strictEqual(ev.env.get('attrnum'), 1, 'FileAttr(..., 2) はVBAファイル番号');
assert.strictEqual(typeof ev.env.get('attrfd'), 'number', 'FileAttr(..., 1) はOSハンドル');
ev.callProcedure('Test4Invalid', []);
assert.strictEqual(ev.env.get('attrerr'), 5, 'FileAttrの無効ReturnTypeはError 5');
console.log('[PASS] FileAttr return type and validation');

let nullError: any = null;
try { evalVBASingle('Debug.Print FileAttr(1, Null)', { fs: new MemoryFileSystem() }); }
catch (error: any) { nullError = error; }
assert.strictEqual(nullError?.number, 94, 'FileAttrのNullはError 94');
console.log('[PASS] FileAttr Null -> Error 94');

// Test 5: Dir attributes distinguish normal files from directories.
ev.callProcedure('Test5DirAttributes', []);
assert.strictEqual(ev.env.get('dirnormal'), '', 'Dir(..., vbNormal) はディレクトリを返さない');
assert.strictEqual(ev.env.get('dirdirectory'), 'test_dir', 'Dir(..., vbDirectory) はディレクトリを含む');
console.log('[PASS] Dir attributes filter directories');

// Test 6: FileCopy must not overwrite an existing destination.
ev.callProcedure('Test6FileCopyExistingDestination', []);
assert.strictEqual(ev.env.get('copyerr'), 58, 'FileCopy existing destination -> Error 58');
console.log('[PASS] FileCopy existing destination');

// Test 7: File-number functions reject Null with Error 94.
ev.callProcedure('Test7FileNumberNull', []);
assert.strictEqual(ev.env.get('eofnullerr'), 94, 'EOF(Null) -> Error 94');
assert.strictEqual(ev.env.get('lofnullerr'), 94, 'LOF(Null) -> Error 94');
assert.strictEqual(ev.env.get('locnullerr'), 94, 'Loc(Null) -> Error 94');
    assert.strictEqual(ev.env.get('seeknullerr'), 94, 'Seek(Null) -> Error 94');
    assert.strictEqual(ev.env.get('opennullerr'), 94, 'Open file number Null -> Error 94');
    console.log('[PASS] File-number Null validation');

// Test 8: MkDir rejects an existing path with Error 75.
ev.callProcedure('Test8MkDirExistingPath', []);
    assert.strictEqual(ev.env.get('mkdirerr'), 75, 'MkDir existing path -> Error 75');
    assert.strictEqual(ev.env.get('mkdirmissingerr'), 76, 'MkDir missing parent -> Error 76');
console.log('[PASS] MkDir existing path');

// Test 9: Put/Get/Seek record positions reject Null with Error 94.
ev.callProcedure('Test9RecordPositionNull', []);
assert.strictEqual(ev.env.get('putrecordnullerr'), 94, 'Put record Null -> Error 94');
assert.strictEqual(ev.env.get('getrecordnullerr'), 94, 'Get record Null -> Error 94');
assert.strictEqual(ev.env.get('seekpositionnullerr'), 94, 'Seek position Null -> Error 94');
console.log('[PASS] Put/Get/Seek record position Null validation');

// Test 10: RmDir rejects non-empty directories with Error 75.
ev.callProcedure('Test10RmDirNonEmpty', []);
assert.strictEqual(ev.env.get('rmdirnonemptyerr'), 75, 'RmDir non-empty directory -> Error 75');
console.log('[PASS] RmDir non-empty directory validation');

// Test 11: FSO OpenTextFile rejects unsupported IOMode values with Error 5.
ev.callProcedure('Test11FsoIOModeBoundary', []);
assert.strictEqual(ev.env.get('fsomodeerr'), 5, 'OpenTextFile IOMode 3 -> Error 5');
console.log('[PASS] FSO OpenTextFile IOMode validation');

// Test 12: Kill rejects directory targets with Error 75.
ev.callProcedure('Test12KillDirectory', []);
assert.strictEqual(ev.env.get('killdirectoryerr'), 75, 'Kill directory target -> Error 75');
console.log('[PASS] Kill directory target validation');

// Test 13: FSO DeleteFolder defaults Force to False for non-empty folders.
ev.callProcedure('Test13FsoDeleteFolderForce', []);
assert.strictEqual(ev.env.get('deletefoldererr'), 75, 'DeleteFolder non-empty without Force -> Error 75');
console.log('[PASS] FSO DeleteFolder Force validation');

// Test 14: Open Output does not create missing parent directories.
ev.callProcedure('Test14OpenOutputMissingParent', []);
assert.strictEqual(ev.env.get('openmissingparenterr'), 76, 'Open Output missing parent -> Error 76');
console.log('[PASS] Open Output parent path validation');

// Test 15: FSO CopyFolder recursively copies a folder tree.
ev.callProcedure('Test15FsoCopyFolder', []);
assert.strictEqual(ev.env.get('copyfolderresult'), 'ok', 'CopyFolder recursive copy');
console.log('[PASS] FSO CopyFolder implementation');

// Test 16: FSO MoveFolder moves a folder tree and removes its source.
ev.callProcedure('Test16FsoMoveFolder', []);
assert.strictEqual(ev.env.get('movefolderresult'), 'ok', 'MoveFolder recursive move');
console.log('[PASS] FSO MoveFolder implementation');

// Test 22: File/Folder path objects share explicit directory/file capabilities.
ev.callProcedure('Test22FsoPathObjectCopyMove', []);
assert.strictEqual(ev.env.get('foldercopyresult'), 'ok', 'GetFolder.Copy recursive copy');
assert.strictEqual(ev.env.get('foldermoveresult'), 'ok', 'GetFolder.Move recursive move');
assert.strictEqual(ev.env.get('filecopyresult'), 'ok', 'GetFile.Copy');
assert.strictEqual(ev.env.get('filemoveresult'), 'ok', 'GetFile.Move');
console.log('[PASS] FSO path object Copy/Move capability contract');

const nodeRoot = nodeFs.mkdtempSync(nodePath.join(nodeOs.tmpdir(), 'vba-fso-path-'));
try {
    const nodeCode = String.raw`
        Public result
        Sub Probe()
            Dim fso As Object, stream As Object, folder As Object
            Set fso = CreateObject("Scripting.FileSystemObject")
            fso.CreateFolder "source"
            Set stream = fso.CreateTextFile("source\child.txt")
            stream.Write "x": stream.Close
            Set folder = fso.GetFolder("source")
            folder.Copy "copy"
            result = IIf(fso.FolderExists("copy") And fso.FileExists("copy\child.txt"), "ok", "bad")
        End Sub
    `;
    const nodeEv = evalVBASingle(nodeCode, { fs: new NodeFileSystem(), sandboxRoot: nodeRoot });
    nodeEv.callProcedure('Probe', []);
    assert.strictEqual(nodeEv.env.get('result'), 'ok', 'NodeFileSystem GetFolder.Copy');
    console.log('[PASS] NodeFileSystem FSO path object Copy');
} finally {
    nodeFs.rmSync(nodeRoot, { recursive: true, force: true });
}

// Test 17: SetAttr rejects vbDirectory as a settable attribute.
ev.callProcedure('Test17SetAttrDirectory', []);
assert.strictEqual(ev.env.get('setattrdirectoryerr'), 5, 'SetAttr vbDirectory -> Error 5');
console.log('[PASS] SetAttr Directory attribute validation');

// Test 18: FSO DeleteFile must honor the Force argument for read-only files.
// Test 18: FSO CreateFolder shares the single-level path contract.
ev.callProcedure('Test18FsoCreateFolderBoundaries', []);
assert.strictEqual(ev.env.get('fsofolderexistingerr'), 58, 'CreateFolder existing path -> Error 58');
assert.strictEqual(ev.env.get('fsofoldermissingparenterr'), 76, 'CreateFolder missing parent -> Error 76');
console.log('[PASS] FSO CreateFolder path validation');

// Test 20: FSO DeleteFile must honor the Force argument for read-only files.
ev.callProcedure('Test20FsoDeleteFileForce', []);
assert.strictEqual(ev.env.get('deletefilenoforceerr'), 70, 'DeleteFile read-only without Force -> Error 70');
assert.strictEqual(ev.env.get('deletefilenoforceexists'), vbaTrue, 'DeleteFile without Force keeps read-only file');
assert.strictEqual(ev.env.get('deletefileforceerr'), 0, 'DeleteFile read-only with Force succeeds');
assert.strictEqual(ev.env.get('deletefileforceexists'), vbaFalse, 'DeleteFile with Force removes read-only file');
console.log('[PASS] FSO DeleteFile Force validation');

// Test 19: FSO DeleteFile must reject directory targets as path access errors.
ev.callProcedure('Test21FsoDeleteFileDirectory', []);
assert.strictEqual(ev.env.get('deletefiledirectoryerr'), 75, 'DeleteFile directory target -> Error 75');
console.log('[PASS] FSO DeleteFile directory validation');

console.log('\n✅ FileSystem (Extra): 全テスト通過');
