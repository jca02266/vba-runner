/**
 * FileSystem Functions (Extra) (§6.1.2.5) のテスト
 */
import { Lexer } from '../../src/compiler/lexer';
import { Parser } from '../../src/compiler/parser';
import { Evaluator } from '../../src/compiler/evaluator';
import { assert } from '../ts/test-runner';
import * as fs from 'fs';
import * as path from 'path';

function evalVBA(code: string): any {
    const tokens = new Lexer(code).tokenize();
    const ast = new Parser(tokens).parse();
    const ev = new Evaluator(console.log);
    ev.evaluate(ast);
    return ev;
}

// Sandbox 用ディレクトリ作成
const sandboxRoot = path.join(process.cwd(), 'sandbox');
if (!fs.existsSync(sandboxRoot)) fs.mkdirSync(sandboxRoot);

// --- 1. Put / Get (Binary mode) ---
const binaryCode = `
    Public s
    Sub Test()
        Open "test_bin.dat" For Binary As #1
        Put #1, , "VBA-Compiler"
        Close #1

        Open "test_bin.dat" For Binary As #2
        Get #2, , s
        Close #2
        
        Debug.Print s
    End Sub
`;

const ev1 = evalVBA(binaryCode);
ev1.callProcedure('Test', []);
assert.strictEqual(ev1.env.get('s'), "VBA-Compiler", 'Put/Get (Binary)');
console.log('[PASS] Put/Get (Binary)');

// --- 2. FileLen, FileDateTime ---
const infoCode = `
    Public flen, fdate
    Sub Test()
        flen = FileLen("test_bin.dat")
        fdate = FileDateTime("test_bin.dat")
    End Sub
`;
const ev2 = evalVBA(infoCode);
ev2.callProcedure('Test', []);
assert.ok(ev2.env.get('flen') > 0, 'FileLen > 0');
assert.ok(ev2.env.get('fdate') && ev2.env.get('fdate').__isVbaDate__, 'FileDateTime');
console.log('[PASS] FileLen, FileDateTime');

// --- 3. Seek, Kill ---
const seekKillCode = `
    Public posBefore, posAfter
    Sub Test()
        Open "test_bin.dat" For Binary As #1
        posBefore = Seek(1)
        Seek #1, 5
        posAfter = Seek(1)
        Close #1
        Kill "test_bin.dat"
    End Sub
`;
const ev3 = evalVBA(seekKillCode);
ev3.callProcedure('Test', []);
assert.strictEqual(ev3.env.get('posbefore'), 1, 'Seek(1) 初期位置');
assert.strictEqual(ev3.env.get('posafter'), 5, 'Seek #1, 5');
assert.strictEqual(fs.existsSync(path.join(sandboxRoot, 'test_bin.dat')), false, 'Kill');
console.log('[PASS] Seek, Kill');

console.log('\n✅ FileSystem (Extra): 全テスト通過');
