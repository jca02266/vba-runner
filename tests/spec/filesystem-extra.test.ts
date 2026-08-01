/**
 * FileSystem Functions (Extra) (§6.1.2.5) のテスト
 */
import { evalVBASingle, assert } from '../../test-libs/test-runner';
import { MemoryFileSystem } from '../../src/engine/filesystem';

// Use VFS (MemoryFileSystem) for tests
const vfs = new MemoryFileSystem();

function evalVBA(code: string): any {
    const ev = evalVBASingle(code, { fs: vfs });
    return ev;
}

// すべてのテストを1つのコード内で実行（複数のプロシージャ）
const allCode = `
    Public s, flen, fdate, posBefore, posAfter, attrFd, attrNum, attrErr, dirNormal, dirDirectory, copyErr
    Public eofNullErr, lofNullErr, locNullErr, seekNullErr, openNullErr, mkdirErr, mkdirMissingErr
    Public putRecordNullErr, getRecordNullErr, seekPositionNullErr

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

console.log('\n✅ FileSystem (Extra): 全テスト通過');
