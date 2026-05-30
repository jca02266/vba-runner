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
import { evalVBAModules, assert } from '../../test-libs/test-runner';

// 1. Module1.bas の Const を Module1.A として参照
{
    const ev = evalVBAModules([
        { name: 'Module1', code: `
Const A = 1
Const B = 42
` },
        { name: 'Module2', code: `
Function Test1() As Long
    Test1 = Module1.A
End Function
Function Test2() As Long
    Test2 = Module1.B
End Function
` },
    ]);

    assert.strictEqual(ev.callProcedure('Test1', []), 1, 'Module1.A = 1');
    console.log('[PASS] Module1.A =', ev.callProcedure('Test1', []));
    assert.strictEqual(ev.callProcedure('Test2', []), 42, 'Module1.B = 42');
    console.log('[PASS] Module1.B =', ev.callProcedure('Test2', []));
}

// 2. Module-level Dim 変数も参照できる
{
    const ev = evalVBAModules([
        { name: 'Module1', code: `
Dim Counter As Long
Counter = 10
` },
        { name: 'Module2', code: `
Function Test3() As Long
    Test3 = Module1.Counter
End Function
` },
    ]);

    assert.strictEqual(ev.callProcedure('Test3', []), 10, 'Module1.Counter = 10');
    console.log('[PASS] Module1.Counter =', ev.callProcedure('Test3', []));
}

// 3. 同名 Const が複数モジュールにある場合は修飾で区別できる
{
    const ev = evalVBAModules([
        { name: 'ModA', code: `Const MAX = 100` },
        { name: 'ModB', code: `Const MAX = 200` },
        { name: 'TestMod', code: `
Function TestA() As Long: TestA = ModA.MAX: End Function
Function TestB() As Long: TestB = ModB.MAX: End Function
` },
    ]);

    assert.strictEqual(ev.callProcedure('TestA', []), 100, 'ModA.MAX = 100');
    assert.strictEqual(ev.callProcedure('TestB', []), 200, 'ModB.MAX = 200');
    console.log('[PASS] ModA.MAX =', ev.callProcedure('TestA', []), 'ModB.MAX =', ev.callProcedure('TestB', []));
}

console.log('\n✅ Module-qualified variable/const access: 全テスト通過');
