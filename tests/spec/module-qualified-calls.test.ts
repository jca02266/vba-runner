/**
 * Module-Qualified Procedure Calls のテスト
 *
 * 複数のモジュール（ファイル）で同じ名前のプロシージャが定義されている場合、
 * ModuleName.ProcedureName() の形式で呼び出し、各モジュール独立してプロシージャを区別して実行する。
 *
 * 網羅する観点:
 *   1. 単一 Evaluator インスタンスに複数モジュールを登録
 *   2. モジュール修飾で異なるモジュールの同名関数を呼び分ける
 *   3. モジュールをまたいだプロシージャ呼び出し
 */
import { Lexer } from '../../src/engine/lexer';
import { Parser } from '../../src/engine/parser';
import { Evaluator } from '../../src/engine/evaluator';
import { assert } from '../../test-libs/test-runner';

function loadAndEvaluate(evaluator: Evaluator, code: string, moduleName: string): void {
    const tokens = new Lexer(code).tokenize();
    const ast = new Parser(tokens).parse();
    evaluator.setSourceModule(moduleName);
    evaluator.evaluate(ast);
}

// --- 1. 複数モジュールを単一 Evaluator に登録し、モジュール修飾で呼び分ける ---
const ev = new Evaluator(console.log);

const moduleACode = `
Function GetValue()
    GetValue = 100
End Function

Function GetLabel()
    GetLabel = "ModuleA"
End Function
`;

const moduleBCode = `
Function GetValue()
    GetValue = 200
End Function

Function GetLabel()
    GetLabel = "ModuleB"
End Function
`;

// 複数モジュールを同一インスタンスに読み込む
loadAndEvaluate(ev, moduleACode, 'ModuleA');
loadAndEvaluate(ev, moduleBCode, 'ModuleB');

// モジュール修飾で各モジュールのプロシージャを呼び分ける
const resultA = ev.callProcedure('ModuleA.GetValue', []);
const resultB = ev.callProcedure('ModuleB.GetValue', []);
const labelA = ev.callProcedure('ModuleA.GetLabel', []);
const labelB = ev.callProcedure('ModuleB.GetLabel', []);

assert.strictEqual(resultA, 100, 'ModuleA.GetValue は 100 を返す');
assert.strictEqual(resultB, 200, 'ModuleB.GetValue は 200 を返す');
assert.strictEqual(labelA, "ModuleA", 'ModuleA.GetLabel は "ModuleA" を返す');
assert.strictEqual(labelB, "ModuleB", 'ModuleB.GetLabel は "ModuleB" を返す');
console.log('[PASS] 複数モジュール間での同名関数の呼び分け');

// --- 2. あるモジュール内のプロシージャが別モジュールのプロシージャを呼び出す ---
const ev2 = new Evaluator(console.log);

const helperCode = `
Function Helper()
    Helper = 42
End Function
`;

const callerCode = `
Function CallHelper()
    ' ModuleB の Helper を ModuleA から呼び出す（モジュール修飾）
    CallHelper = ModuleB.Helper()
End Function
`;

loadAndEvaluate(ev2, helperCode, 'ModuleB');
loadAndEvaluate(ev2, callerCode, 'ModuleA');

const crossModuleResult = ev2.callProcedure('ModuleA.CallHelper', []);
assert.strictEqual(crossModuleResult, 42, 'ModuleA から ModuleB のプロシージャを呼び出し');
console.log('[PASS] モジュールをまたいだプロシージャ呼び出し');

// --- 3. モジュール内のプロシージャは修飾あり/なしどちらでも呼び出し可能 ---
const ev3 = new Evaluator(console.log);

const globalCode = `
Function GlobalFunc()
    GlobalFunc = 999
End Function
`;

const moduleZCode = `
Function ModuleFunc()
    ModuleFunc = 333
End Function
`;

// グローバルコード（Module1 として登録）
loadAndEvaluate(ev3, globalCode, 'Module1');
// モジュール名付きコード
loadAndEvaluate(ev3, moduleZCode, 'ModuleZ');

// グローバル関数は修飾なしで呼び出し可能
const globalResult = ev3.callProcedure('GlobalFunc', []);
assert.strictEqual(globalResult, 999, 'グローバル関数は修飾なしで呼び出し可能');

// モジュール関数は修飾付きで呼び出し可能
const moduleResultQualified = ev3.callProcedure('ModuleZ.ModuleFunc', []);
assert.strictEqual(moduleResultQualified, 333, 'モジュール関数は修飾付きで呼び出し可能');

// モジュール関数は修飾なしでも呼び出し可能（唯一の同名関数の場合）
const moduleResultUnqualified = ev3.callProcedure('ModuleFunc', []);
assert.strictEqual(moduleResultUnqualified, 333, 'モジュール関数は修飾なしでも呼び出し可能');

console.log('[PASS] モジュール内プロシージャは修飾あり/なしで呼び出し可能');

console.log('[PASS] モジュール修飾付きプロシージャ呼び出し - 全テスト完了');
