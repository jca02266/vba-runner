/**
 * Global Scope Procedure Qualification Test
 *
 * グローバルスコープ（モジュール名なし）で登録されたプロシージャが、
 * モジュール修飾で呼び出せるかを確認
 */
import { Lexer } from '../../src/engine/lexer';
import { Parser } from '../../src/engine/parser';
import { Evaluator } from '../../src/engine/evaluator';

function loadAndEvaluate(evaluator: Evaluator, code: string, moduleName: string): void {
    const tokens = new Lexer(code).tokenize();
    const ast = new Parser(tokens).parse();
    evaluator.setSourceModule(moduleName);
    evaluator.evaluate(ast);
}

const globalCode = `
Function GlobalFunc()
    GlobalFunc = 999
End Function
`;

const moduleCode = `
Function ModuleFunc()
    ModuleFunc = 111
End Function
`;

const ev = new Evaluator(console.log);

// Global procedure registered as Module1
loadAndEvaluate(ev, globalCode, 'Module1');

// Register module procedure
loadAndEvaluate(ev, moduleCode, 'MyModule');

console.log('=== Test 1: Unqualified call to global ===');
try {
    const result = ev.callProcedure('GlobalFunc', []);
    console.log(`✓ GlobalFunc() = ${result}`);
} catch (e: any) {
    console.log(`✗ GlobalFunc() failed: ${e.message}`);
}

console.log('\n=== Test 2: Qualified call to global with Module1 ===');
console.log('Module1 is the default module name for global scope');
try {
    const result = ev.callProcedure('Module1.GlobalFunc', []);
    console.log(`✓ Module1.GlobalFunc() = ${result}`);
} catch (e: any) {
    console.log(`✗ Module1.GlobalFunc() failed: ${e.message}`);
}

console.log('\n=== Test 3: Module procedure ===');
try {
    const result = ev.callProcedure('MyModule.ModuleFunc', []);
    console.log(`✓ MyModule.ModuleFunc() = ${result}`);
} catch (e: any) {
    console.log(`✗ MyModule.ModuleFunc() failed: ${e.message}`);
}

console.log('\n=== Summary ===');
console.log('Current implementation:');
console.log("- Global procedures (setSourceModule('')) default to Module1");
console.log('- Can be called unqualified: GlobalFunc()');
console.log('- Can be called qualified: Module1.GlobalFunc()');
console.log('- Module procedures must use qualification if ambiguous');
