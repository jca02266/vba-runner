/**
 * Array Functions (UBound, LBound, Array) のテスト
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

// --- 1. UBound, LBound (1次元) ---
const arrayCode = `
    Public arr(1 To 5) As Integer
    Public lb, ub
    Sub Test()
        lb = LBound(arr)
        ub = UBound(arr)
    End Sub
`;
const ev1 = evalVBA(arrayCode);
ev1.callProcedure('Test', []);
assert.strictEqual(ev1.env.get('lb'), 1, 'LBound(arr) は 1');
assert.strictEqual(ev1.env.get('ub'), 5, 'UBound(arr) は 5');
console.log('[PASS] UBound, LBound (1次元)');

// --- 2. UBound, LBound (多次元) ---
const multiArrayCode = `
    Public arr(1 To 3, 0 To 10) As Integer
    Public lb1, ub1, lb2, ub2
    Sub Test()
        lb1 = LBound(arr, 1)
        ub1 = UBound(arr, 1)
        lb2 = LBound(arr, 2)
        ub2 = UBound(arr, 2)
    End Sub
`;
const ev2 = evalVBA(multiArrayCode);
ev2.callProcedure('Test', []);
assert.strictEqual(ev2.env.get('lb1'), 1, 'LBound(arr, 1)');
assert.strictEqual(ev2.env.get('ub1'), 3, 'UBound(arr, 1)');
assert.strictEqual(ev2.env.get('lb2'), 0, 'LBound(arr, 2)');
assert.strictEqual(ev2.env.get('ub2'), 10, 'UBound(arr, 2)');
console.log('[PASS] UBound, LBound (多次元)');

// --- 3. Array 関数 ---
const arrayFnCode = `
    Public v
    Sub Test()
        v = Array("A", "B", "C")
    End Sub
`;
const ev3 = evalVBA(arrayFnCode);
ev3.callProcedure('Test', []);
const v3 = ev3.env.get('v');
assert.strictEqual(Array.isArray(v3), true, 'Array() は配列を返す');
assert.strictEqual(v3.length, 3, '要素数は 3');
assert.strictEqual(v3[0], "A", 'v(0) は "A"');
console.log('[PASS] Array 関数');

console.log('\n✅ Array Functions: 全テスト通過');
