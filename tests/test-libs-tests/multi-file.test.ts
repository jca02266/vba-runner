/**
 * VBARunner マルチファイルロードのテスト
 *
 * tests/fixtures/multi-file/ に .vba / .cls / .frm の3ファイルを配置し、
 * ディレクトリを指定するだけで全ファイルが同一スコープにロードされ
 * 相互参照できることを検証する。
 */
import { VBARunner, assert } from '../../test-libs/test-runner';

const DIR = 'tests/fixtures/multi-file';

const suite = new VBARunner(DIR);

// --- 1. .vba ファイルのプロシージャが呼べる ---
assert.strictEqual(suite.run('Add', [3, 4]), 7, 'Add (MathUtils.vba)');
assert.strictEqual(suite.run('Multiply', [6, 7]), 42, 'Multiply (MathUtils.vba)');
console.log('[PASS] .vba ファイルのプロシージャ');

// --- 2. .cls ファイルのプロシージャが呼べる ---
assert.strictEqual(suite.run('Greet', ['World']), 'Hello, World!', 'Greet (StringUtils.bas)');
assert.strictEqual(suite.run('Repeat', ['ab', 3]), 'ababab', 'Repeat (StringUtils.bas)');
console.log('[PASS] .cls ファイルのプロシージャ');

// --- 3. .frm ファイルから他ファイルのプロシージャを相互参照できる ---
assert.strictEqual(
    suite.run('CalcAndGreet', ['Alice', 10, 20]),
    'Hello, Alice! Sum=30',
    'CalcAndGreet が Add/Greet を相互参照 (Main.frm)'
);
console.log('[PASS] ファイル間の相互参照');

// --- 4. 単一ファイル指定（後方互換）も引き続き動作する ---
const single = new VBARunner('tests/fixtures/multi-file/MathUtils.vba');
assert.strictEqual(single.run('Add', [1, 2]), 3, '単一ファイル指定の後方互換');
console.log('[PASS] 単一ファイル指定（後方互換）');

console.log('\n✅ マルチファイルロード: 全テスト通過');
