/**
 * FileSystem Functions (Extra) (§6.1.2.5) のテスト
 */
import { Lexer } from '../../src/engine/lexer';
import { Parser } from '../../src/engine/parser';
import { Evaluator } from '../../src/engine/evaluator';
import { assert } from '../../test-libs/test-runner';
import { MemoryFileSystem } from '../../src/engine/filesystem';

// Use VFS (MemoryFileSystem) for tests
const vfs = new MemoryFileSystem();

function evalVBA(code: string): any {
    const tokens = new Lexer(code).tokenize();
    const ast = new Parser(tokens).parse();
    const ev = new Evaluator(console.log, { fs: vfs });
    ev.evaluate(ast);
    return ev;
}

// すべてのテストを1つのコード内で実行（複数のプロシージャ）
const allCode = `
    Public s, flen, fdate, posBefore, posAfter

    Sub Test1()
        ' --- 1. Put / Get (Binary mode) ---
        Open "test_bin.dat" For Binary As #1
        Put #1, , "VBA-Compiler"
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
`;

const ev = evalVBA(allCode);

// Test 1: Put/Get
ev.callProcedure('Test1', []);
assert.strictEqual(ev.env.get('s'), "VBA-Compiler", 'Put/Get (Binary)');
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

console.log('\n✅ FileSystem (Extra): 全テスト通過');
