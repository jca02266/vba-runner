/**
 * Evaluator Scoping Test
 *
 * 複数 Evaluator インスタンスの場合と、単一 Evaluator インスタンスの場合で
 * モジュール修飾呼び出しがどう動作するかを比較
 */
import { Lexer } from '../../src/engine/lexer';
import { Parser } from '../../src/engine/parser';
import { Evaluator } from '../../src/engine/evaluator';

function loadAndEvaluate(evaluator: Evaluator, code: string, moduleName: string): void {
    const tokens = new Lexer(code).tokenize();
    const ast = new Parser(tokens).parse();
    evaluator.setSourceModule(moduleName);
    evaluator.evaluateModule(ast);
    evaluator.resolveIdentifiers([{ ast, moduleName }]);
}

const moduleACode = `
Function GetValue()
    GetValue = 111
End Function
`;

const moduleBCode = `
Function GetValue()
    GetValue = 222
End Function
`;

console.log('=== Test 1: Single Evaluator with Multiple Modules ===');
const ev1 = new Evaluator(console.log);
loadAndEvaluate(ev1, moduleACode, 'ModuleA');
loadAndEvaluate(ev1, moduleBCode, 'ModuleB');

try {
    const resultA = ev1.callProcedure('ModuleA.GetValue', []);
    console.log(`✓ ModuleA.GetValue() = ${resultA}`);
} catch (e: any) {
    console.log(`✗ ModuleA.GetValue() failed: ${e.message}`);
}

try {
    const resultB = ev1.callProcedure('ModuleB.GetValue', []);
    console.log(`✓ ModuleB.GetValue() = ${resultB}`);
} catch (e: any) {
    console.log(`✗ ModuleB.GetValue() failed: ${e.message}`);
}

try {
    const resultUnqualified = ev1.callProcedure('GetValue', []);
    console.log(`✗ GetValue() (unqualified) = ${resultUnqualified} - Should raise ambiguity error!`);
} catch (e: any) {
    console.log(`✓ GetValue() (unqualified) correctly raised: ${e.message.split('\n')[0]}`);
}

console.log('\n=== Test 2: Separate Evaluator Instances ===');
const evA = new Evaluator(console.log);
loadAndEvaluate(evA, moduleACode, 'ModuleA');

const evB = new Evaluator(console.log);
loadAndEvaluate(evB, moduleBCode, 'ModuleB');

console.log('ModuleA (Evaluator A):');
try {
    const resultA = evA.callProcedure('ModuleA.GetValue', []);
    console.log(`✓ ModuleA.GetValue() = ${resultA}`);
} catch (e: any) {
    console.log(`✗ ModuleA.GetValue() failed: ${e.message}`);
}

try {
    const resultModuleB = evA.callProcedure('ModuleB.GetValue', []);
    console.log(`✗ ModuleB.GetValue() found = ${resultModuleB} - ModuleB not in this Evaluator!`);
} catch (e: any) {
    console.log(`✓ ModuleB.GetValue() correctly failed: Procedure not found (as expected)`);
}

console.log('\nModuleB (Evaluator B):');
try {
    const resultB = evB.callProcedure('ModuleB.GetValue', []);
    console.log(`✓ ModuleB.GetValue() = ${resultB}`);
} catch (e: any) {
    console.log(`✗ ModuleB.GetValue() failed: ${e.message}`);
}

try {
    const resultModuleA = evB.callProcedure('ModuleA.GetValue', []);
    console.log(`✗ ModuleA.GetValue() found = ${resultModuleA} - ModuleA not in this Evaluator!`);
} catch (e: any) {
    console.log(`✓ ModuleA.GetValue() correctly failed: Procedure not found (as expected)`);
}

console.log('\n=== Summary ===');
console.log('Single Evaluator:');
console.log('  - All modules share same procedure namespace');
console.log('  - Module qualification works for cross-module calls');
console.log('  - Ambiguity detection works across all modules');
console.log('\nSeparate Evaluators:');
console.log('  - Each Evaluator is isolated');
console.log('  - No cross-module calls possible');
console.log('  - Modules cannot interact');
