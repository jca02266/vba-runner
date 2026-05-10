/**
 * Private スコープのクロスモジュール制御テスト
 *
 * 検証する挙動:
 *   1. テストランナー (vbaTest.run) からは Private プロシージャを直接呼べる
 *   2. 同一モジュール内からは Private プロシージャを呼べる
 *   3. 別モジュールの VBA ソースから Private プロシージャを呼ぶとエラーになる
 *   4. Public プロシージャはどのモジュールからでも呼べる
 */
import { VBATest, assert } from '../ts/test-runner';

const DIR = 'tests/fixtures/scope-private';
const suite = new VBATest(DIR);

// --- 1. テストランナーからは Private を直接呼べる ---
assert.strictEqual(suite.run('PrivateHelper', [5]), 10, 'テストランナーから Private を直接呼べる');
console.log('[PASS] テストランナーから Private を直接呼べる');

// --- 2. 同一モジュール内から Private を呼べる ---
assert.strictEqual(suite.run('UseOwnPrivate', [3]), 6, '同一モジュール内で Private を呼べる (3*2=6)');
console.log('[PASS] 同一モジュール内から Private を呼べる');

// --- 3. Public プロシージャはクロスモジュールで呼べる ---
assert.strictEqual(suite.run('CallPublic', [10, 5]), 15, 'クロスモジュールで Public を呼べる');
console.log('[PASS] クロスモジュールで Public を呼べる');

// --- 4. VBA ソースから別モジュールの Private を呼ぶとエラー ---
let threw = false;
try {
    suite.run('CallPrivateCrossModule', [7]);
} catch (e: any) {
    threw = true;
    assert.strictEqual(
        e.message.includes("Cannot call Private procedure 'PrivateHelper'"),
        true,
        'エラーメッセージに Private プロシージャ名が含まれる'
    );
    assert.strictEqual(
        e.message.includes('ModuleA.vba'),
        true,
        'エラーメッセージに定義元モジュールが含まれる'
    );
    assert.strictEqual(
        e.message.includes('ModuleB.vba'),
        true,
        'エラーメッセージに呼び出し元モジュールが含まれる'
    );
}
assert.strictEqual(threw, true, '別モジュールから Private 呼び出しで例外が発生する');
console.log('[PASS] 別モジュールの Private 呼び出しはエラーになる');

console.log('\n✅ Private スコープ制御: 全テスト通過');
