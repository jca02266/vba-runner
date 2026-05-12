/**
 * ReDim, Erase (§5.4.3.3, §5.4.3.4) のテスト
 */
import { Lexer } from '../../src/compiler/lexer';
import { Parser } from '../../src/compiler/parser';
import { Evaluator } from '../../src/compiler/evaluator';
import { assert } from '../ts/test-runner';

function evalVBA(code: string): any {
    const tokens = new Lexer(code).tokenize();
    const ast = new Parser(tokens).parse();
    const ev = new Evaluator(console.log);
    ev.evaluate(ast);
    return ev;
}

// --- 1. ReDim (動的配列の再定義) ---
const redimCode = `
    Public arr() As Integer
    Sub Test()
        ReDim arr(2)
        arr(0) = 10
        arr(1) = 20
        arr(2) = 30
    End Sub
`;
const ev1 = evalVBA(redimCode);
ev1.callProcedure('Test', []);
const arr1 = ev1.env.get('arr');
assert.strictEqual(Array.isArray(arr1), true, 'arr は配列である');
assert.strictEqual(arr1.length, 3, 'arr の長さは 3 (0 to 2)');
assert.strictEqual(arr1[0], 10, 'arr(0) は 10');
console.log('[PASS] ReDim 基本動作');

// --- 2. ReDim Preserve (データ維持) ---
const preserveCode = `
    Public arr() As Integer
    Sub Test()
        ReDim arr(1)
        arr(0) = 100
        arr(1) = 200
        ReDim Preserve arr(2)
        arr(2) = 300
    End Sub
`;
const ev2 = evalVBA(preserveCode);
ev2.callProcedure('Test', []);
const arr2 = ev2.env.get('arr');
assert.strictEqual(arr2.length, 3, 'arr の長さは 3');
assert.strictEqual(arr2[0], 100, 'データが維持されている (arr(0))');
assert.strictEqual(arr2[1], 200, 'データが維持されている (arr(1))');
assert.strictEqual(arr2[2], 300, '新しい要素がセットされている (arr(2))');
console.log('[PASS] ReDim Preserve');

// --- 3. Erase (動的配列) ---
const eraseDynCode = `
    Public arr() As Integer
    Sub Test()
        ReDim arr(2)
        Erase arr
    End Sub
`;
const ev3 = evalVBA(eraseDynCode);
ev3.callProcedure('Test', []);
const arr3 = ev3.env.get('arr');
// VBA では動的配列を Erase すると未割り当て状態になる
// 現在の実装では [] になる
assert.strictEqual(Array.isArray(arr3), true, 'Erase 後も配列オブジェクト自体は存在する (実装依存)');
assert.strictEqual(arr3.length, 0, '動的配列を Erase すると空になる');
console.log('[PASS] Erase (動的配列)');

// --- 4. Erase (固定配列) ---
const eraseFixedCode = `
    Public arr(2) As Integer
    Sub Test()
        arr(0) = 10
        arr(1) = 20
        arr(2) = 30
        Erase arr
    End Sub
`;
const ev4 = evalVBA(eraseFixedCode);
ev4.callProcedure('Test', []);
const arr4 = ev4.env.get('arr');
// VBA では固定配列を Erase すると全要素が 0 (Integer の場合) に戻る
// 現在の実装では [] になってしまうため、このテストは失敗する可能性がある
assert.strictEqual(arr4.length, 3, '固定配列を Erase してもサイズは変わらない');
assert.strictEqual(arr4[0], 0, '固定配列を Erase すると要素が初期化される (arr(0))');
console.log('[PASS] Erase (固定配列)');

console.log('\n✅ ReDim, Erase: 全テスト通過');
