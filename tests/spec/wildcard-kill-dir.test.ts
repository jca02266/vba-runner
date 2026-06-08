import { MemoryFileSystem } from '../../src/engine/filesystem';
import { evalVBASingle, assert } from '../../test-libs/test-runner';

function runVBA(code: string, vfs: MemoryFileSystem) {
    return evalVBASingle(code, { onPrint: () => {}, fs: vfs });
}

function seedFiles(vfs: MemoryFileSystem, files: string[]) {
    for (const f of files) {
        vfs.writeFileSync(f, 'test');
    }
}

// Test 1: Kill with *.ext wildcard deletes matching files
{
    const vfs = new MemoryFileSystem();
    seedFiles(vfs, ['/sandbox/a.txt', '/sandbox/b.txt', '/sandbox/c.log', '/sandbox/d.txt']);

    runVBA(`Kill "*.txt"`, vfs);

    assert.strictEqual(vfs.existsSync('/sandbox/a.txt'), false, 'a.txt deleted');
    assert.strictEqual(vfs.existsSync('/sandbox/b.txt'), false, 'b.txt deleted');
    assert.strictEqual(vfs.existsSync('/sandbox/d.txt'), false, 'd.txt deleted');
    assert.strictEqual(vfs.existsSync('/sandbox/c.log'), true, 'c.log preserved');
    console.log('[PASS] Kill *.txt wildcard');
}

// Test 2: Kill with ? wildcard (exactly one character)
{
    const vfs = new MemoryFileSystem();
    seedFiles(vfs, ['/sandbox/file1.dat', '/sandbox/file2.dat', '/sandbox/file12.dat', '/sandbox/other.dat']);

    runVBA(`Kill "file?.dat"`, vfs);

    assert.strictEqual(vfs.existsSync('/sandbox/file1.dat'), false, 'file1.dat deleted');
    assert.strictEqual(vfs.existsSync('/sandbox/file2.dat'), false, 'file2.dat deleted');
    assert.strictEqual(vfs.existsSync('/sandbox/file12.dat'), true, 'file12.dat (2 chars) preserved');
    assert.strictEqual(vfs.existsSync('/sandbox/other.dat'), true, 'other.dat preserved');
    console.log('[PASS] Kill file?.dat wildcard');
}

// Test 3: Kill *.* deletes all files in directory (not subdirs)
{
    const vfs = new MemoryFileSystem();
    seedFiles(vfs, ['/sandbox/a.txt', '/sandbox/b.log', '/sandbox/c.tmp']);

    runVBA(`Kill "*.*"`, vfs);

    assert.strictEqual(vfs.existsSync('/sandbox/a.txt'), false, 'a.txt deleted');
    assert.strictEqual(vfs.existsSync('/sandbox/b.log'), false, 'b.log deleted');
    assert.strictEqual(vfs.existsSync('/sandbox/c.tmp'), false, 'c.tmp deleted');
    console.log('[PASS] Kill *.* wildcard');
}

// Test 4: Kill with no-match pattern raises Error 53 (caught by On Error Resume Next in a function)
{
    const vfs = new MemoryFileSystem();
    seedFiles(vfs, ['/sandbox/readme.txt']);

    const code = `
        Function TestNoMatch() As Integer
            On Error Resume Next
            Kill "*.xyz"
            TestNoMatch = Err.Number
        End Function
    `;
    const errNum = evalVBASingle(code, { onPrint: () => {}, fs: vfs }).callProcedure('TestNoMatch', []);
    assert.strictEqual(errNum, 53, 'Kill no-match raises Error 53');
    // Source file is preserved
    assert.strictEqual(vfs.existsSync('/sandbox/readme.txt'), true, 'readme.txt still exists');
    console.log('[PASS] Kill no-match pattern raises Error 53');
}


// Test 5: Kill exact path (no wildcard) still works
{
    const vfs = new MemoryFileSystem();
    seedFiles(vfs, ['/sandbox/exact.txt']);

    runVBA(`Kill "exact.txt"`, vfs);

    assert.strictEqual(vfs.existsSync('/sandbox/exact.txt'), false, 'exact.txt deleted');
    console.log('[PASS] Kill exact path (no wildcard)');
}

// Test 6: Dir with * wildcard lists matching files
{
    const vfs = new MemoryFileSystem();
    seedFiles(vfs, ['/sandbox/alpha.txt', '/sandbox/beta.txt', '/sandbox/gamma.log']);

    const code = `
        Function CollectDir() As String
            Dim result As String
            Dim f As String
            f = Dir("*.txt")
            Do While f <> ""
                result = result & f & ","
                f = Dir()
            Loop
            CollectDir = result
        End Function
    `;
    const result = String(evalVBASingle(code, { onPrint: () => {}, fs: vfs }).callProcedure('CollectDir', []));
    const parts = result.split(',').filter(s => s !== '');
    assert.strictEqual(parts.includes('alpha.txt'), true, 'Dir *.txt includes alpha.txt');
    assert.strictEqual(parts.includes('beta.txt'), true, 'Dir *.txt includes beta.txt');
    assert.strictEqual(parts.includes('gamma.log'), false, 'Dir *.txt excludes gamma.log');
    console.log('[PASS] Dir *.txt wildcard');
}

// Test 7: Dir with ? wildcard
{
    const vfs = new MemoryFileSystem();
    seedFiles(vfs, ['/sandbox/file1.dat', '/sandbox/file2.dat', '/sandbox/file12.dat']);

    const code = `
        Function CollectDir() As String
            Dim result As String
            Dim f As String
            f = Dir("file?.dat")
            Do While f <> ""
                result = result & f & ","
                f = Dir()
            Loop
            CollectDir = result
        End Function
    `;
    const result = String(evalVBASingle(code, { onPrint: () => {}, fs: vfs }).callProcedure('CollectDir', []));
    const parts = result.split(',').filter(s => s !== '');
    assert.strictEqual(parts.includes('file1.dat'), true, 'Dir file?.dat includes file1.dat');
    assert.strictEqual(parts.includes('file2.dat'), true, 'Dir file?.dat includes file2.dat');
    assert.strictEqual(parts.includes('file12.dat'), false, 'Dir file?.dat excludes file12.dat (2 chars)');
    console.log('[PASS] Dir file?.dat wildcard');
}

// Test 8: Real-world pattern - bulk delete temp files before processing
{
    const vfs = new MemoryFileSystem();
    seedFiles(vfs, [
        '/sandbox/report_2024.tmp',
        '/sandbox/report_2025.tmp',
        '/sandbox/report_2024.xlsx',
        '/sandbox/notes.txt'
    ]);

    const code = `
        Sub CleanupTempFiles()
            On Error Resume Next
            Kill "*.tmp"
        End Sub
        CleanupTempFiles
    `;
    runVBA(code, vfs);

    assert.strictEqual(vfs.existsSync('/sandbox/report_2024.tmp'), false, '.tmp files deleted');
    assert.strictEqual(vfs.existsSync('/sandbox/report_2025.tmp'), false, '.tmp files deleted');
    assert.strictEqual(vfs.existsSync('/sandbox/report_2024.xlsx'), true, '.xlsx preserved');
    assert.strictEqual(vfs.existsSync('/sandbox/notes.txt'), true, '.txt preserved');
    console.log('[PASS] Real-world: bulk delete *.tmp files');
}

// Test 9: Real-world pattern - collect files matching pattern with Dir loop
{
    const vfs = new MemoryFileSystem();
    seedFiles(vfs, [
        '/sandbox/data_2024_01.csv',
        '/sandbox/data_2024_02.csv',
        '/sandbox/data_2024_03.csv',
        '/sandbox/summary.xlsx'
    ]);

    const code = `
        Function CountCsvFiles() As Integer
            Dim count As Integer
            Dim f As String
            count = 0
            f = Dir("data_*.csv")
            Do While f <> ""
                count = count + 1
                f = Dir()
            Loop
            CountCsvFiles = count
        End Function
    `;
    const count = evalVBASingle(code, { onPrint: () => {}, fs: vfs }).callProcedure('CountCsvFiles', []);
    assert.strictEqual(count, 3, 'Dir loop counts 3 CSV files');
    console.log('[PASS] Real-world: Dir loop counting files');
}

console.log('\n✅ Wildcard Kill/Dir: 全テスト通過');
