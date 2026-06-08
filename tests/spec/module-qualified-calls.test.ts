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
import { evalVBAModules, assert } from '../../test-libs/test-runner';

function loadAndEvaluate(evaluator: Evaluator, code: string, moduleName: string): void {
    const tokens = new Lexer(code).tokenize();
    const ast = new Parser(tokens).parse();
    evaluator.setSourceModule(moduleName);
    evaluator.evaluateModule(ast);
    evaluator.resolveIdentifiers([{ ast, moduleName }]);
}

// --- 1. 複数モジュールを単一 Evaluator に登録し、モジュール修飾で呼び分ける ---
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

const ev = evalVBAModules([
    { name: 'ModuleA', code: moduleACode },
    { name: 'ModuleB', code: moduleBCode },
]);

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
const helperCode = `
Function Helper()
    Helper = 42
End Function
`;

const callerCode = `
Function CallHelper()
    CallHelper = ModuleB.Helper()
End Function
`;

const ev2 = evalVBAModules([
    { name: 'ModuleB', code: helperCode },
    { name: 'ModuleA', code: callerCode },
]);

const crossModuleResult = ev2.callProcedure('ModuleA.CallHelper', []);
assert.strictEqual(crossModuleResult, 42, 'ModuleA から ModuleB のプロシージャを呼び出し');
console.log('[PASS] モジュールをまたいだプロシージャ呼び出し');

// --- 3. モジュール内のプロシージャは修飾あり/なしどちらでも呼び出し可能 ---
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

const ev3 = evalVBAModules([
    { name: 'Module1', code: globalCode },
    { name: 'ModuleZ', code: moduleZCode },
]);

const globalResult = ev3.callProcedure('GlobalFunc', []);
assert.strictEqual(globalResult, 999, 'グローバル関数は修飾なしで呼び出し可能');

const moduleResultQualified = ev3.callProcedure('ModuleZ.ModuleFunc', []);
assert.strictEqual(moduleResultQualified, 333, 'モジュール関数は修飾付きで呼び出し可能');

const moduleResultUnqualified = ev3.callProcedure('ModuleFunc', []);
assert.strictEqual(moduleResultUnqualified, 333, 'モジュール関数は修飾なしでも呼び出し可能');

console.log('[PASS] モジュール内プロシージャは修飾あり/なしで呼び出し可能');

// --- 4. Call Module.ProcName（引数なし）のパースが成功する ---
{
    const src = `
Sub Test()
    Call Module1.DoSomething
End Sub
`;
    const tokens = new Lexer(src).tokenize();
    const ast = new Parser(tokens).parse();
    const diags = ast.diagnostics ?? [];
    assert.strictEqual(diags.length, 0, 'Call Module.ProcName はパースエラーにならない');
    const stmt = (ast.body[0] as any).body[0];
    assert.strictEqual(stmt.type, 'CallStatement', 'CallStatement として解析される');
    assert.strictEqual(stmt.expression.type, 'CallExpression', 'expression が CallExpression');
    assert.strictEqual(stmt.expression.callee.type, 'MemberExpression', 'callee が MemberExpression');
    console.log('[PASS] Call Module.ProcName のパース');
}

// --- 5. Call Module.ProcName(args) のパースが成功する ---
{
    const src = `
Sub Test()
    Call Module1.DoSomething(1, 2)
End Sub
`;
    const tokens = new Lexer(src).tokenize();
    const ast = new Parser(tokens).parse();
    const diags = ast.diagnostics ?? [];
    assert.strictEqual(diags.length, 0, 'Call Module.ProcName(args) はパースエラーにならない');
    const stmt = (ast.body[0] as any).body[0];
    assert.strictEqual(stmt.expression.args.length, 2, '引数が2つ認識される');
    console.log('[PASS] Call Module.ProcName(args) のパース');
}

console.log('[PASS] モジュール修飾付きプロシージャ呼び出し - 全テスト完了');
