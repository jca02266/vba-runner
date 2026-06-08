/**
 * AddressOf Expression (§5.6.16.8) のテスト
 */
import { evalVBASingle, assert } from '../../test-libs/test-runner';

function evalVBA(code: string): any {
    return evalVBASingle(code);
}

// --- 1. AddressOf の基本 ---
const addrCode = `
    Public addr
    Sub MyProc()
    End Sub
    Sub Test()
        addr = AddressOf MyProc
    End Sub
`;
const ev1 = evalVBA(addrCode);
ev1.callProcedure('Test', []);
// 現在の実装では関数名を文字列として返す
assert.strictEqual(ev1.env.get('addr'), "MyProc", 'AddressOf MyProc');
console.log('[PASS] AddressOf 基本動作');

console.log('\n✅ AddressOf: 全テスト通過');
