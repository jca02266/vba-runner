/**
 * Ambiguous Procedure Call Resolution Test
 *
 * VBA Name Resolution Priorities (修飾なし呼び出し):
 *   1. Unqualified procedures (global scope)
 *   2. Single module-qualified procedure (unambiguous)
 *   3. Multiple module-qualified procedures → Error (ambiguous)
 *
 * Test three scenarios:
 *   1. Global procedure takes priority
 *   2. Single module procedure works
 *   3. Multiple modules → Ambiguity error (requires qualification)
 */
import { Lexer } from '../../src/compiler/lexer';
import { Parser } from '../../src/compiler/parser';
import { Evaluator } from '../../src/compiler/evaluator';
import { assert } from '../../test-libs/test-runner';

function loadAndEvaluate(evaluator: Evaluator, code: string, moduleName: string): void {
    const tokens = new Lexer(code).tokenize();
    const ast = new Parser(tokens).parse();
    evaluator.setSourceModule(moduleName);
    evaluator.evaluate(ast);
}

// --- Test 1: Global procedure (Module1) accessible both ways ---
const ev1 = new Evaluator(console.log);

const globalCode = `
Function GetValue()
    GetValue = 999
End Function
`;

const moduleCalcCode = `
Function Calculate()
    Calculate = 555
End Function
`;

loadAndEvaluate(ev1, globalCode, 'Module1');  // Registered as Module1
loadAndEvaluate(ev1, moduleCalcCode, 'Module2');

// Unqualified call to global procedure
const result1a = ev1.callProcedure('GetValue', []);
assert.strictEqual(result1a, 999, 'Global procedure accessible unqualified');

// Qualified call to global procedure via Module1
const result1b = ev1.callProcedure('Module1.GetValue', []);
assert.strictEqual(result1b, 999, 'Global procedure accessible as Module1.GetValue');

console.log('[PASS] Global procedure (Module1) accessible both ways');

// --- Test 2: Single module procedure works with unqualified call ---
const ev2 = new Evaluator(console.log);

const module2Code = `
Function Calculate()
    Calculate = 555
End Function
`;

loadAndEvaluate(ev2, module2Code, 'Module2');

const result2 = ev2.callProcedure('Calculate', []);
assert.strictEqual(result2, 555, 'Single module procedure accessible without qualification');
console.log('[PASS] Single module procedure unqualified call works');

// --- Test 3: Ambiguity error when multiple modules have same procedure ---
const ev3 = new Evaluator(console.log);

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

loadAndEvaluate(ev3, moduleACode, 'ModuleA');
loadAndEvaluate(ev3, moduleBCode, 'ModuleB');

// Unqualified call should throw ambiguity error
let ambiguityErrorThrown = false;
try {
    ev3.callProcedure('GetValue', []);
} catch (e: any) {
    if (e.message.includes('Ambiguous procedure')) {
        ambiguityErrorThrown = true;
    }
}
assert.ok(ambiguityErrorThrown, 'Ambiguity error thrown for multiple module procedures');
console.log('[PASS] Ambiguity error correctly raised');

// Qualified calls must be used to disambiguate
const resultA = ev3.callProcedure('ModuleA.GetValue', []);
const resultB = ev3.callProcedure('ModuleB.GetValue', []);
assert.strictEqual(resultA, 111, 'ModuleA.GetValue returns 111');
assert.strictEqual(resultB, 222, 'ModuleB.GetValue returns 222');
console.log('[PASS] Qualified calls disambiguate correctly');

console.log('[PASS] VBA Name Resolution Priority - all tests passed');
