/**
 * Private スコープのクロスモジュール制御テスト (§5.2.3.1 / §5.3.1.1)
 *
 * 検証する挙動:
 *   1. テストランナー (callProcedure) からは Private プロシージャを直接呼べる
 *   2. 同一モジュール内からは Private プロシージャを呼べる
 *   3. 別モジュールの VBA ソースから Private プロシージャを呼ぶとエラーになる
 *   4. Public プロシージャはどのモジュールからでも呼べる
 */
import { Lexer } from '../../src/engine/lexer';
import { Parser } from '../../src/engine/parser';
import { Evaluator } from '../../src/engine/evaluator';
import { assert } from '../../test-libs/test-runner';

const moduleA = `
    Public Function PublicAdd(a, b)
        PublicAdd = a + b
    End Function

    Private Function PrivateHelper(x)
        PrivateHelper = x * 2
    End Function

    Public Function UseOwnPrivate(x)
        UseOwnPrivate = PrivateHelper(x)
    End Function
`;

const moduleB = `
    Public Function CallPublic(a, b)
        CallPublic = PublicAdd(a, b)
    End Function

    Public Function CallPrivateCrossModule(x)
        CallPrivateCrossModule = PrivateHelper(x)
    End Function
`;

// 2 つのモジュールを別々にセットして評価
const ev = new Evaluator(console.log);

ev.setSourceModule('ModuleA.vba');
ev.evaluate(new Parser(new Lexer(moduleA).tokenize()).parse());

ev.setSourceModule('ModuleB.vba');
ev.evaluate(new Parser(new Lexer(moduleB).tokenize()).parse());

ev.setSourceModule('');

console.log('--- Starting Private Scope Cross-Module Tests ---');

// --- 1. テストランナーからは Private を直接呼べる ---
assert.strictEqual(ev.callProcedure('PrivateHelper', [5]), 10, 'テストランナーから Private を直接呼べる (5*2=10)');
console.log('[PASS] テストランナーから Private を直接呼べる');

// --- 2. 同一モジュール内から Private を呼べる ---
assert.strictEqual(ev.callProcedure('UseOwnPrivate', [3]), 6, '同一モジュール内で Private を呼べる (3*2=6)');
console.log('[PASS] 同一モジュール内から Private を呼べる');

// --- 3. Public プロシージャはクロスモジュールで呼べる ---
assert.strictEqual(ev.callProcedure('CallPublic', [10, 5]), 15, 'クロスモジュールで Public を呼べる (10+5=15)');
console.log('[PASS] クロスモジュールで Public を呼べる');

// --- 4. VBA ソースから別モジュールの Private を呼ぶとエラー ---
let threw = false;
try {
    ev.callProcedure('CallPrivateCrossModule', [7]);
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
