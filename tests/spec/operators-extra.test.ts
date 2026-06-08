/**
 * Operators (Extra) (§5.6.9) のテスト
 */
import { evalVBASingle, assert } from '../../test-libs/test-runner';

function evalVBA(code: string): any {
    return evalVBASingle(code);
}

// --- 1. 文字列連結 (&) ---
const concatCode = `
    Public s
    Sub Test()
        s = "Hello" & " " & 123 & True
    End Sub
`;
const ev1 = evalVBA(concatCode);
ev1.callProcedure('Test', []);
assert.strictEqual(ev1.env.get('s'), "Hello 123True", '文字列連結 (&)');
console.log('[PASS] 文字列連結 (&)');

// --- 2. 比較演算子 (=, <>, <, >, <=, >=) ---
const compCode = `
    Public eq, ne, lt, gt, le, ge
    Sub Test()
        eq = (10 = 10)
        ne = (10 <> 20)
        lt = (10 < 20)
        gt = (20 > 10)
        le = (10 <= 10)
        ge = (10 >= 10)
    End Sub
`;
const ev2 = evalVBA(compCode);
ev2.callProcedure('Test', []);
assert.strictEqual(ev2.env.get('eq').value, -1, '10 = 10');
assert.strictEqual(ev2.env.get('ne').value, -1, '10 <> 20');
assert.strictEqual(ev2.env.get('lt').value, -1, '10 < 20');
assert.strictEqual(ev2.env.get('gt').value, -1, '20 > 10');
assert.strictEqual(ev2.env.get('le').value, -1, '10 <= 10');
assert.strictEqual(ev2.env.get('ge').value, -1, '10 >= 10');
console.log('[PASS] 比較演算子');

// --- 3. Like 演算子 ---
const likeCode = `
    Public res1, res2, res3, res4
    Sub Test()
        res1 = ("abcde" Like "a*e")
        res2 = ("abcde" Like "a?c?e")
        res3 = ("abcde" Like "a#cde") ' Fail: # は数字 [0-9]
        res4 = ("abcde" Like "[a-z]bcde")
    End Sub
`;
const ev3 = evalVBA(likeCode);
ev3.callProcedure('Test', []);
assert.strictEqual(ev3.env.get('res1').value, -1, '"abcde" Like "a*e"');
assert.strictEqual(ev3.env.get('res2').value, -1, '"abcde" Like "a?c?e"');
assert.strictEqual(ev3.env.get('res3').value, 0, '"abcde" Like "a#cde" (False)');
assert.strictEqual(ev3.env.get('res4').value, -1, '"abcde" Like "[a-z]bcde"');
console.log('[PASS] Like 演算子');

console.log('\n✅ Operators (Extra): 全テスト通過');
