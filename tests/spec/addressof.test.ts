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

// --- 2. AddressOf Module.Proc (§5.6.16.8 メンバー形式) ---
{
    const code = `
        Public addr
        Sub Test()
            addr = AddressOf MyModule.MyProc
        End Sub
    `;
    const ev = evalVBA(code);
    ev.callProcedure('Test', []);
    assert.strictEqual(ev.env.get('addr'), 'MyModule.MyProc', 'AddressOf Module.Proc → "Module.Proc"');
    console.log('[PASS] AddressOf Module.Proc 形式');
}

// --- 3. AddressOf でコンテキストキーワード関数名もモジュール修飾可能 ---
{
    const code = `
        Public addr
        Sub Test()
            addr = AddressOf Utils.Step
        End Sub
    `;
    const ev = evalVBA(code);
    ev.callProcedure('Test', []);
    assert.strictEqual(ev.env.get('addr'), 'Utils.Step', 'AddressOf Module.contextualKw');
    console.log('[PASS] AddressOf Module.ContextualKeyword 形式');
}

console.log('\n✅ AddressOf: 全テスト通過');
