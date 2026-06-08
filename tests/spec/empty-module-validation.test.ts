/**
 * Empty Module Name Validation Test
 *
 * 空のモジュール名による登録を防止
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

const simpleCode = `
Function GetValue()
    GetValue = 123
End Function
`;

console.log('=== Test 1: Empty module name on fresh evaluator ===');
const ev1 = new Evaluator(console.log);

try {
    loadAndEvaluate(ev1, simpleCode, '');
    console.log('✗ Empty module name was accepted (should have failed)');
} catch (e: any) {
    console.log(`✓ Empty module name rejected: ${e.message.split('\n')[0]}`);
}

console.log('\n=== Test 2: Explicit module name works ===');
const ev2 = new Evaluator(console.log);

try {
    loadAndEvaluate(ev2, simpleCode, 'MyModule');
    const result = ev2.callProcedure('MyModule.GetValue', []);
    console.log(`✓ Explicit module name 'MyModule' works: result = ${result}`);
} catch (e: any) {
    console.log(`✗ Explicit module name failed: ${e.message}`);
}

console.log('\n=== Test 3: Attribute VB_Name sets module name ===');
const ev3 = new Evaluator(console.log);

const codeWithAttribute = `
Attribute VB_Name = "ProperModule"

Function Calculate()
    Calculate = 456
End Function
`;

try {
    const tokens = new Lexer(codeWithAttribute).tokenize();
    const ast = new Parser(tokens).parse();
    ev3.evaluateModule(ast);
    ev3.resolveIdentifiers([{ ast, moduleName: '' }]);
    const result = ev3.callProcedure('ProperModule.Calculate', []);
    console.log(`✓ Attribute VB_Name module name works: result = ${result}`);
} catch (e: any) {
    console.log(`✗ Attribute VB_Name failed: ${e.message}`);
}

console.log('\n=== Test 4: Reset module name with explicit value ===');
const ev4 = new Evaluator(console.log);

try {
    loadAndEvaluate(ev4, simpleCode, 'Module1');
    ev4.setSourceModule('');
    const result = ev4.callProcedure('Module1.GetValue', []);
    console.log(`✓ setSourceModule('') preserves previous module: result = ${result}`);
} catch (e: any) {
    console.log(`✗ Module preservation failed: ${e.message}`);
}

console.log('\n=== Summary ===');
console.log('Empty module name validation:');
console.log('- Cannot set empty module on fresh Evaluator');
console.log('- Must provide explicit module name OR');
console.log('- Use Attribute VB_Name in VBA source');
console.log("- Once a module is set, setSourceModule('') preserves it");
