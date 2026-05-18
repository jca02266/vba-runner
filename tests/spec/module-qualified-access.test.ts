/**
 * Module-qualified variable/constant access: Module1.A style reference
 *
 * Verifies that constants and module-level variables defined in one module
 * can be referenced from another module using the Module.Name syntax.
 *
 * Implementation notes:
 * - Constants: stored with module-qualified key (e.g. "module1:a") — enables
 *   same-name disambiguation across modules since constants are immutable.
 * - Variables: registered in moduleVarRegistry; lookup uses unqualified name —
 *   same-name variables across modules share the global slot (last-write wins),
 *   so disambiguation of same-name variables is a known limitation.
 */
import { Lexer } from '../../src/engine/lexer';
import { Parser } from '../../src/engine/parser';
import { Evaluator } from '../../src/engine/evaluator';
import { VBARunner, assert } from '../../test-libs/test-runner';

// 1. Module1.bas の Const を Module1.A として参照
{
    const vr = new VBARunner();
    const mod1 = `
Const A = 1
Const B = 42
`;
    const ast1 = new Parser(new Lexer(mod1).tokenize()).parse();
    vr.evaluator.setSourceModule('Module1');
    vr.evaluator.evaluate(ast1);

    const mod2 = `
Function Test1() As Long
    Test1 = Module1.A
End Function
Function Test2() As Long
    Test2 = Module1.B
End Function
`;
    const ast2 = new Parser(new Lexer(mod2).tokenize()).parse();
    vr.evaluator.setSourceModule('Module2');
    vr.evaluator.evaluate(ast2);

    assert.strictEqual(vr.evaluator.callProcedure('Test1', []), 1, 'Module1.A = 1');
    console.log('[PASS] Module1.A =', vr.evaluator.callProcedure('Test1', []));
    assert.strictEqual(vr.evaluator.callProcedure('Test2', []), 42, 'Module1.B = 42');
    console.log('[PASS] Module1.B =', vr.evaluator.callProcedure('Test2', []));
}

// 2. Module-level Dim 変数も参照できる
{
    const vr = new VBARunner();
    const mod1 = `
Dim Counter As Long
Counter = 10
`;
    const ast1 = new Parser(new Lexer(mod1).tokenize()).parse();
    vr.evaluator.setSourceModule('Module1');
    vr.evaluator.evaluate(ast1);

    const mod2 = `
Function Test3() As Long
    Test3 = Module1.Counter
End Function
`;
    const ast2 = new Parser(new Lexer(mod2).tokenize()).parse();
    vr.evaluator.setSourceModule('Module2');
    vr.evaluator.evaluate(ast2);

    assert.strictEqual(vr.evaluator.callProcedure('Test3', []), 10, 'Module1.Counter = 10');
    console.log('[PASS] Module1.Counter =', vr.evaluator.callProcedure('Test3', []));
}

// 3. 同名 Const が複数モジュールにある場合は修飾で区別できる
{
    const vr = new VBARunner();
    const mod1 = `Const MAX = 100`;
    const mod2 = `Const MAX = 200`;

    const ast1 = new Parser(new Lexer(mod1).tokenize()).parse();
    vr.evaluator.setSourceModule('ModA');
    vr.evaluator.evaluate(ast1);

    const ast2 = new Parser(new Lexer(mod2).tokenize()).parse();
    vr.evaluator.setSourceModule('ModB');
    vr.evaluator.evaluate(ast2);

    const ast3 = new Parser(new Lexer(`
Function TestA() As Long: TestA = ModA.MAX: End Function
Function TestB() As Long: TestB = ModB.MAX: End Function
`).tokenize()).parse();
    vr.evaluator.setSourceModule('TestMod');
    vr.evaluator.evaluate(ast3);

    assert.strictEqual(vr.evaluator.callProcedure('TestA', []), 100, 'ModA.MAX = 100');
    assert.strictEqual(vr.evaluator.callProcedure('TestB', []), 200, 'ModB.MAX = 200');
    console.log('[PASS] ModA.MAX =', vr.evaluator.callProcedure('TestA', []), 'ModB.MAX =', vr.evaluator.callProcedure('TestB', []));
}

console.log('\n✅ Module-qualified variable/const access: 全テスト通過');
