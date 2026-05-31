import { Lexer } from '../../src/engine/lexer';
import { Parser } from '../../src/engine/parser';
import { Evaluator } from '../../src/engine/evaluator';
import { assert, evalVBASingle } from '../../test-libs/test-runner';

// Option Explicit チェックは Pass 2（resolveIdentifiers）に一本化されているため、
// 単一モジュールのテストは Pass 1+2 を両方実行する evalVBASingle を使う。
function evalVBA(code: string): any {
    return evalVBASingle(code);
}

function runFunc(code: string, name: string, args: any[] = []): any {
    return evalVBA(code).callProcedure(name, args);
}

// --- 1. No Option Explicit: undeclared variables are fine ---
{
    const code = `
Function Test() As String
    x = 42
    Test = CStr(x)
End Function
`;
    const result = runFunc(code, 'Test');
    assert.strictEqual(result, '42', 'No Option Explicit: implicit variable works');
    console.log('[PASS] No Option Explicit: implicit variable allowed');
}

// --- 2. Option Explicit: declared variable works ---
{
    const code = `
Option Explicit
Function Test() As String
    Dim x As Long
    x = 42
    Test = CStr(x)
End Function
`;
    const result = runFunc(code, 'Test');
    assert.strictEqual(result, '42', 'Option Explicit with Dim: works');
    console.log('[PASS] Option Explicit: declared variable works');
}

// --- 3. Option Explicit: undeclared variable in Sub causes runtime error when called ---
{
    const code = `
Option Explicit
Sub BadSub()
    x = 5
End Sub
`;
    let threw = false;
    try {
        runFunc(code, 'BadSub');
    } catch (e: any) {
        threw = true;
        assert.strictEqual(typeof e.message === 'string' && e.message.includes('not declared'), true, 'Error mentions not declared');
    }
    assert.strictEqual(threw, true, 'Option Explicit: undeclared variable throws at call time');
    console.log('[PASS] Option Explicit: undeclared variable throws runtime error');
}

// --- 4. Valid Sub can still run even when another Sub has violations ---
{
    const code = `
Option Explicit

Function GoodFunc() As Long
    Dim n As Long
    n = 99
    GoodFunc = n
End Function

Sub BadSub()
    x = 5
End Sub
`;
    const ev = evalVBA(code);

    // GoodFunc should work fine
    const goodResult = ev.callProcedure('GoodFunc', []);
    assert.strictEqual(goodResult, 99, 'Good function still runs despite BadSub violation');

    // BadSub should throw
    let threw = false;
    try {
        ev.callProcedure('BadSub', []);
    } catch {
        threw = true;
    }
    assert.strictEqual(threw, true, 'BadSub with undeclared var throws');
    console.log('[PASS] Option Explicit: good subs run; bad sub throws');
}

// --- 5. Parameters are implicitly declared ---
{
    const code = `
Option Explicit
Function AddTwo(a As Long, b As Long) As Long
    AddTwo = a + b
End Function
`;
    const result = runFunc(code, 'AddTwo', [3, 4]);
    assert.strictEqual(result, 7, 'Parameters are implicitly declared under Option Explicit');
    console.log('[PASS] Option Explicit: parameters are valid without Dim');
}

// --- 6. For loop variable requires Dim ---
{
    const code = `
Option Explicit
Function Test() As Long
    Dim total As Long
    For i = 1 To 3
        total = total + i
    Next i
    Test = total
End Function
`;
    let threw = false;
    try {
        runFunc(code, 'Test');
    } catch {
        threw = true;
    }
    assert.strictEqual(threw, true, 'For loop variable without Dim throws under Option Explicit');
    console.log('[PASS] Option Explicit: For loop var requires Dim');
}

// --- 7. For loop variable with Dim works ---
{
    const code = `
Option Explicit
Function Test() As Long
    Dim total As Long
    Dim i As Long
    For i = 1 To 3
        total = total + i
    Next i
    Test = total
End Function
`;
    const result = runFunc(code, 'Test');
    assert.strictEqual(result, 6, 'For loop with Dim i works');
    console.log('[PASS] Option Explicit: For loop var with Dim works');
}

// --- 8. Dim after use is a violation (no hoisting) ---
{
    const code = `
Option Explicit
Sub BadSub()
    b = 5
    Dim b As Long
End Sub
`;
    let threw = false;
    try {
        runFunc(code, 'BadSub');
    } catch {
        threw = true;
    }
    assert.strictEqual(threw, true, 'Dim after use is an error (no hoisting)');
    console.log('[PASS] Option Explicit: Dim after use is an error');
}

// --- 9. Module-level Dim accessible in all Subs ---
{
    const code = `
Option Explicit
Dim moduleVar As Long

Sub SetIt()
    moduleVar = 42
End Sub

Function GetIt() As Long
    GetIt = moduleVar
End Function
`;
    const ev = evalVBA(code);
    ev.callProcedure('SetIt', []);
    const result = ev.callProcedure('GetIt', []);
    assert.strictEqual(result, 42, 'Module-level Dim accessible in Subs');
    console.log('[PASS] Option Explicit: module-level Dim accessible in Subs');
}

// --- 10. Diagnostics are populated for static analysis ---
{
    const tokens = new Lexer(`
Option Explicit
Sub BadSub()
    x = 5
End Sub
`).tokenize();
    const ast = new Parser(tokens).parse();
    const ev = new Evaluator(console.log);
    // Option Explicit の静的解析は Pass 2（resolveIdentifiers）で実行される。
    ev.evaluateModule(ast);
    ev.resolveIdentifiers([{ ast, moduleName: '' }]);
    const errors = ast.diagnostics.filter(d => d.message.includes('not declared'));
    assert.strictEqual(errors.length > 0, true, 'Diagnostics contain undeclared variable error');
    assert.strictEqual(errors[0].severity, 'error', 'Diagnostic severity is error');
    console.log('[PASS] Option Explicit: diagnostics populated for static analysis');
}

// --- 11. Multi-module: undeclared object in call expression is flagged in 2nd pass ---
// evalVBAModules を使うことで reEvaluateModuleConstsAll が全モジュール名を持ち、
// bare identifier が本当に未宣言かどうかを精密に判定できる。
{
    const { evalVBAModules } = await import('../../test-libs/test-runner');
    const ev = evalVBAModules([
        {
            name: 'ModA',
            code: `
Option Explicit
Sub Test()
    undeclaredObj.Method()
End Sub
`,
        },
    ]);
    let threw = false;
    try {
        ev.callProcedure('Test', []);
    } catch (e: any) {
        threw = true;
        assert.strictEqual(e.message.includes('not declared') || e.message.includes('undeclaredObj'), true,
            'Error mentions undeclaredObj: ' + e.message);
    }
    assert.strictEqual(threw, true, 'undeclaredObj.Method() flagged when module names are known');
    console.log('[PASS] Option Explicit: undeclaredObj.Method() detected in multi-module 2nd pass');
}

// --- 12. Multi-module: known module name in call is NOT flagged ---
{
    const { evalVBAModules } = await import('../../test-libs/test-runner');
    const ev = evalVBAModules([
        {
            name: 'Helper',
            code: `
Public Function Add(a As Long, b As Long) As Long
    Add = a + b
End Function
`,
        },
        {
            name: 'Main',
            code: `
Option Explicit
Function RunTest() As Long
    RunTest = Helper.Add(3, 4)
End Function
`,
        },
    ]);
    const result = ev.callProcedure('RunTest', []);
    assert.strictEqual(result, 7, 'Known module name Helper.Add() is not flagged');
    console.log('[PASS] Option Explicit: known module name in call is not flagged');
}

console.log('\n✅ Option Explicit: 全テスト通過');
