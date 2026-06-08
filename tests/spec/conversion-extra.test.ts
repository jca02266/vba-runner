/**
 * Conversion Functions (Extra) (§6.1.2.3) のテスト
 */
import { evalVBASingle, assert } from '../../test-libs/test-runner';

function evalVBA(code: string): any {
    return evalVBASingle(code);
}

// --- 1. CBool, CInt, CLng, CDbl, CStr ---
const convCode = `
    Public b, i, l, d, s
    Sub Test()
        b = CBool(1)
        i = CInt(12.5) ' VBA CInt は銀行型丸め (一番近い偶数へ) -> 現在の実装は Math.round (四捨五入)
        l = CLng(123456.7)
        d = CDbl("123.45")
        s = CStr(123)
    End Sub
`;
const ev1 = evalVBA(convCode);
ev1.callProcedure('Test', []);
assert.strictEqual(ev1.env.get('b').value, -1, 'CBool(1) は True');
// 12.5 -> 12 (偶数丸め) or 13 (四捨五入)
// vbaRound を使っているはず
assert.strictEqual(ev1.env.get('i'), 12, 'CInt(12.5) は 12 (銀行型丸め)');
assert.strictEqual(ev1.env.get('l'), 123457, 'CLng(123456.7) は 123457');
assert.strictEqual(ev1.env.get('d'), 123.45, 'CDbl("123.45")');
assert.strictEqual(ev1.env.get('s'), "123", 'CStr(123)');
console.log('[PASS] 基本変換関数');

// --- 2. Int, Val ---
const intValCode = `
    Public v1, v2, v3, v4
    Sub Test()
        v1 = Int(-8.4) ' -9
        v2 = Int(8.4)  ' 8
        v3 = Val(" 123.45 abc")
        v4 = Val("1,234.56") ' Val はカンマを無視するかどうか (VBA 本来は無視しないが現在の実装は置換している)
    End Sub
`;
const ev2 = evalVBA(intValCode);
ev2.callProcedure('Test', []);
assert.strictEqual(ev2.env.get('v1'), -9, 'Int(-8.4) は -9');
assert.strictEqual(ev2.env.get('v2'), 8, 'Int(8.4) は 8');
assert.strictEqual(ev2.env.get('v3'), 123.45, 'Val(" 123.45 abc")');
console.log('[PASS] Int, Val');

console.log('\n✅ Conversion Functions (Extra): 全テスト通過');
