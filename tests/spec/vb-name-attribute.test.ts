/**
 * Attribute VB_Name Module Definition Test
 *
 * VBA ソースファイルが Attribute VB_Name で定義したモジュール名が
 * 正しく認識されるかを確認
 */
import { Lexer } from '../../src/engine/lexer';
import { Parser } from '../../src/engine/parser';
import { Evaluator } from '../../src/engine/evaluator';
import { assert } from '../../test-libs/test-runner';

function loadAndEvaluate(evaluator: Evaluator, code: string): void {
    const tokens = new Lexer(code).tokenize();
    const ast = new Parser(tokens).parse();
    evaluator.evaluateModule(ast);
    evaluator.resolveIdentifiers([{ ast, moduleName: '' }]);
}

// --- Test 1: Attribute VB_Name determines module name ---
const ev1 = new Evaluator(console.log);

const vbaWithAttribute = `
Attribute VB_Name = "MyCustomModule"

Function GetValue()
    GetValue = 111
End Function
`;

loadAndEvaluate(ev1, vbaWithAttribute);

try {
    const result = ev1.callProcedure('MyCustomModule.GetValue', []);
    assert.strictEqual(result, 111, 'Attribute VB_Name module name is recognized');
    console.log('[PASS] Attribute VB_Name module name recognized');
} catch (e: any) {
    console.log(`[FAIL] MyCustomModule.GetValue failed: ${e.message}`);
    throw e;
}

// --- Test 2: Multiple modules with Attribute VB_Name ---
const ev2 = new Evaluator(console.log);

const moduleA = `
Attribute VB_Name = "ExportedModuleA"

Function Calculate()
    Calculate = 222
End Function
`;

const moduleB = `
Attribute VB_Name = "ExportedModuleB"

Function Calculate()
    Calculate = 333
End Function
`;

loadAndEvaluate(ev2, moduleA);
loadAndEvaluate(ev2, moduleB);

const resultA = ev2.callProcedure('ExportedModuleA.Calculate', []);
const resultB = ev2.callProcedure('ExportedModuleB.Calculate', []);

assert.strictEqual(resultA, 222, 'ExportedModuleA.Calculate');
assert.strictEqual(resultB, 333, 'ExportedModuleB.Calculate');

console.log('[PASS] Multiple modules with Attribute VB_Name work correctly');

// --- Test 3: Unqualified call with ambiguous modules ---
let ambiguityErrorThrown = false;
try {
    ev2.callProcedure('Calculate', []);
} catch (e: any) {
    if (e.message.includes('Ambiguous')) {
        ambiguityErrorThrown = true;
    }
}
assert.ok(ambiguityErrorThrown, 'Ambiguity error raised for multiple modules');
console.log('[PASS] Ambiguity detection works with Attribute VB_Name');

console.log('[PASS] Attribute VB_Name - all tests passed');
