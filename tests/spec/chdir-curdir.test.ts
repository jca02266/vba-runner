import { Lexer } from '../../src/engine/lexer';
import { Parser } from '../../src/engine/parser';
import { Evaluator } from '../../src/engine/evaluator';
import { MemoryFileSystem } from '../../src/engine/filesystem';
import { assert } from '../../test-libs/test-runner';

function makeEval(vfs: MemoryFileSystem): Evaluator {
    return new Evaluator(() => {}, { fs: vfs });
}

function runVBA(code: string, vfs: MemoryFileSystem): any {
    const tokens = new Lexer(code).tokenize();
    const ast = new Parser(tokens).parse();
    const ev = new Evaluator(() => {}, { fs: vfs });
    ev.evaluateModule(ast);
    ev.resolveIdentifiers([{ ast, moduleName: '' }]);
    return ev;
}

function runFunc(code: string, name: string, vfs: MemoryFileSystem): any {
    const tokens = new Lexer(code).tokenize();
    const ast = new Parser(tokens).parse();
    const ev = new Evaluator(() => {}, { fs: vfs });
    ev.evaluateModule(ast);
    ev.resolveIdentifiers([{ ast, moduleName: '' }]);
    return ev.callProcedure(name, []);
}

// Test 1: CurDir() initially returns the sandbox root
{
    const vfs = new MemoryFileSystem();
    const result = runFunc(`
        Function Test1() As String
            Test1 = CurDir()
        End Function
    `, 'Test1', vfs);
    // CurDir() returns a VBA-style path; root maps to "\" in virtual path
    assert.strictEqual(typeof result, 'string', 'CurDir returns a string');
    assert.strictEqual(result.length > 0, true, 'CurDir returns non-empty string');
    console.log('[PASS] CurDir() returns initial directory:', result);
}

// Test 2: ChDir changes the directory; CurDir reflects it
{
    const vfs = new MemoryFileSystem();
    vfs.mkdirSync('/sandbox/subdir', { recursive: true });

    const result = runFunc(`
        Function Test2() As String
            ChDir "subdir"
            Test2 = CurDir()
        End Function
    `, 'Test2', vfs);
    assert.strictEqual(result.includes('subdir'), true, 'CurDir reflects ChDir change');
    console.log('[PASS] ChDir changes CurDir:', result);
}

// Test 3: ChDir ".." navigates up one level
{
    const vfs = new MemoryFileSystem();
    vfs.mkdirSync('/sandbox/alpha/beta', { recursive: true });

    const result = runFunc(`
        Function Test3() As String
            Dim d1 As String, d2 As String
            ChDir "alpha"
            d1 = CurDir()
            ChDir ".."
            d2 = CurDir()
            Test3 = d1 & "|" & d2
        End Function
    `, 'Test3', vfs);
    const parts = String(result).split('|');
    assert.strictEqual(parts[0].includes('alpha'), true, 'After ChDir "alpha", CurDir contains "alpha"');
    assert.strictEqual(parts[1].includes('alpha'), false, 'After ChDir "..", CurDir no longer contains "alpha"');
    console.log('[PASS] ChDir ".." navigates up:', result);
}

// Test 4: Relative file Open uses cwd after ChDir
{
    const vfs = new MemoryFileSystem();
    vfs.mkdirSync('/sandbox/data', { recursive: true });
    vfs.writeFileSync('/sandbox/data/hello.txt', 'hello');

    const result = runFunc(`
        Function Test4() As String
            Dim f As Integer
            Dim s As String
            ChDir "data"
            f = FreeFile
            Open "hello.txt" For Input As #f
            Line Input #f, s
            Close #f
            Test4 = s
        End Function
    `, 'Test4', vfs);
    assert.strictEqual(result, 'hello', 'Open uses cwd after ChDir');
    console.log('[PASS] Relative Open uses cwd after ChDir');
}

// Test 5: Kill with wildcard uses cwd after ChDir
{
    const vfs = new MemoryFileSystem();
    vfs.mkdirSync('/sandbox/tmp', { recursive: true });
    vfs.writeFileSync('/sandbox/tmp/a.tmp', 'x');
    vfs.writeFileSync('/sandbox/tmp/b.tmp', 'x');
    vfs.writeFileSync('/sandbox/keep.txt', 'x');

    runVBA(`
        ChDir "tmp"
        Kill "*.tmp"
    `, vfs);

    assert.strictEqual(vfs.existsSync('/sandbox/tmp/a.tmp'), false, 'a.tmp deleted via cwd');
    assert.strictEqual(vfs.existsSync('/sandbox/tmp/b.tmp'), false, 'b.tmp deleted via cwd');
    assert.strictEqual(vfs.existsSync('/sandbox/keep.txt'), true, 'keep.txt untouched');
    console.log('[PASS] Kill wildcard uses cwd after ChDir');
}

// Test 6: Dir() uses cwd after ChDir
{
    const vfs = new MemoryFileSystem();
    vfs.mkdirSync('/sandbox/docs', { recursive: true });
    vfs.writeFileSync('/sandbox/docs/readme.txt', 'r');
    vfs.writeFileSync('/sandbox/docs/notes.txt', 'n');
    vfs.writeFileSync('/sandbox/other.txt', 'o');

    const result = runFunc(`
        Function Test6() As String
            Dim f As String, r As String
            ChDir "docs"
            f = Dir("*.txt")
            Do While f <> ""
                r = r & f & ","
                f = Dir()
            Loop
            Test6 = r
        End Function
    `, 'Test6', vfs);
    const parts = String(result).split(',').filter(s => s !== '');
    assert.strictEqual(parts.includes('readme.txt'), true, 'Dir finds readme.txt in cwd');
    assert.strictEqual(parts.includes('notes.txt'), true, 'Dir finds notes.txt in cwd');
    assert.strictEqual(parts.includes('other.txt'), false, 'Dir excludes file outside cwd');
    console.log('[PASS] Dir() uses cwd after ChDir');
}

// Test 7: ChDir with nested subdirectory (multiple levels)
{
    const vfs = new MemoryFileSystem();
    vfs.mkdirSync('/sandbox/foo/bar/baz', { recursive: true });

    const result = runFunc(`
        Function Test7() As String
            ChDir "foo"
            ChDir "bar"
            ChDir "baz"
            Test7 = CurDir()
        End Function
    `, 'Test7', vfs);
    assert.strictEqual(String(result).includes('baz'), true, 'Nested ChDir descends correctly');
    console.log('[PASS] Nested ChDir:', result);
}

// Test 8: Real-world pattern — process files in a subdirectory then return
{
    const vfs = new MemoryFileSystem();
    vfs.mkdirSync('/sandbox/reports', { recursive: true });
    vfs.writeFileSync('/sandbox/reports/jan.csv', 'jan');
    vfs.writeFileSync('/sandbox/reports/feb.csv', 'feb');

    const result = runFunc(`
        Function Test8() As Integer
            Dim count As Integer
            Dim f As String
            ChDir "reports"
            f = Dir("*.csv")
            Do While f <> ""
                count = count + 1
                f = Dir()
            Loop
            ChDir ".."
            Test8 = count
        End Function
    `, 'Test8', vfs);
    assert.strictEqual(result, 2, 'Dir loop counts files in subdirectory');
    console.log('[PASS] Real-world: process files in subdir then return');
}

console.log('\n✅ ChDir/CurDir: 全テスト通過');
