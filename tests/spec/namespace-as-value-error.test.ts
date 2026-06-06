/**
 * プロジェクト名・モジュール名を値として使った場合のエラー検出テスト
 *
 * VBA 仕様:
 * - VarType(VBA)       → コンパイルエラー「プロジェクトではなく、変数またはプロシージャを指定してください。」
 * - VarType(Scripting) → 同上（型ライブラリ名もプロジェクト扱い）
 * - VarType(Mod1)      → コンパイルエラー「モジュールではなく、変数またはプロシージャを指定してください。」
 * 修飾子としての VBA.X・Scripting.Dictionary・Mod1.Proc は引き続き正常。
 */
import { evalVBASingle, evalVBAModules, assert } from '../../test-libs/test-runner';
import { Evaluator } from '../../src/engine/evaluator';
import { Lexer } from '../../src/engine/lexer';
import { Parser } from '../../src/engine/parser';
import type { VbaComObject } from '../../src/engine/vba-types';

function catchError(modules: Array<{ name: string; code: string }>, entry: string): any {
    const ev = evalVBAModules(modules);
    try {
        ev.callProcedure(entry, []);
    } catch (e) {
        return e;
    }
    assert.fail(`expected '${entry}' to throw`);
    throw new Error('unreachable');
}

function noError(modules: Array<{ name: string; code: string }>, entry: string): void {
    const ev = evalVBAModules(modules);
    try {
        ev.callProcedure(entry, []);
    } catch (e: any) {
        assert.fail(`expected '${entry}' not to throw, but got: ${e.message}`);
    }
}

// ---------------------------------------------------------------------------
// VBA（プロジェクト名）を値として使う → エラー
// ---------------------------------------------------------------------------
{
    const e = catchError([{ name: 'Main', code: `
Function Test() As Long
    Test = VarType(VBA)
End Function`}], 'Test');
    assert.ok(e, 'VarType(VBA) はエラーを投げる');
    assert.ok(e.message.includes('project') || e.message.includes('プロジェクト'), `エラーメッセージに "project" を含む: "${e.message}"`);
    console.log('[PASS] VarType(VBA) → エラー');
}
{
    const e = catchError([{ name: 'Main', code: `
Sub Test()
    Dim x As Long
    x = VBA
End Sub`}], 'Test');
    assert.ok(e, 'x = VBA はエラーを投げる');
    console.log('[PASS] x = VBA → エラー');
}

// ---------------------------------------------------------------------------
// モジュール名を値として使う → エラー
// ---------------------------------------------------------------------------
{
    const e = catchError([
        { name: 'Mod1', code: `Public Function Foo() As Long : Foo = 1 : End Function` },
        { name: 'Main', code: `
Function Test() As Long
    Test = VarType(Mod1)
End Function`}
    ], 'Test');
    assert.ok(e, 'VarType(Mod1) はエラーを投げる');
    assert.ok(e.message.includes('module') || e.message.includes('モジュール'), `エラーメッセージに "module" を含む: "${e.message}"`);
    console.log('[PASS] VarType(Mod1) → エラー');
}
{
    const e = catchError([
        { name: 'Mod1', code: `Public Function Foo() As Long : Foo = 1 : End Function` },
        { name: 'Main', code: `
Sub Test()
    Dim x As Long
    x = Mod1
End Sub`}
    ], 'Test');
    assert.ok(e, 'x = Mod1 はエラーを投げる');
    console.log('[PASS] x = Mod1 → エラー');
}

// ---------------------------------------------------------------------------
// 修飾子としての VBA.X は引き続き正常
// ---------------------------------------------------------------------------
noError([{ name: 'Main', code: `
Function Test() As Long
    Test = VBA.VarType("hello")
End Function`}], 'Test');
console.log('[PASS] VBA.VarType(...) → 正常');

noError([{ name: 'Main', code: `
Function Test() As Long
    Test = VBA.vbString
End Function`}], 'Test');
console.log('[PASS] VBA.vbString → 正常');

// ---------------------------------------------------------------------------
// 修飾子としてのモジュール名.プロシージャ は正常
// ---------------------------------------------------------------------------
noError([
    { name: 'Mod1', code: `Public Function Foo() As Long : Foo = 42 : End Function` },
    { name: 'Main', code: `
Function Test() As Long
    Test = Mod1.Foo()
End Function`}
], 'Test');
console.log('[PASS] Mod1.Foo() → 正常');

// ---------------------------------------------------------------------------
// VBA と同名の変数を宣言したらその変数として使える（エラーなし）
// ---------------------------------------------------------------------------
noError([{ name: 'Main', code: `
Function Test() As Long
    Dim VBA As Long
    VBA = 5
    Test = VBA
End Function`}], 'Test');
console.log('[PASS] Dim VBA As Long → 変数として正常使用');

// ---------------------------------------------------------------------------
// 組み込み型ライブラリ名（Scripting 等）もプロジェクト扱いでエラー
// ---------------------------------------------------------------------------
{
    const e = catchError([{ name: 'Main', code: `
Function Test() As Long
    Test = VarType(Scripting)
End Function`}], 'Test');
    assert.ok(e, 'VarType(Scripting) はエラーを投げる');
    assert.ok(e.message.includes('project') || e.message.includes('プロジェクト'), `エラーメッセージに "project" を含む: "${e.message}"`);
    console.log('[PASS] VarType(Scripting) → エラー（型ライブラリ名はプロジェクト扱い）');
}

// Scripting.Dictionary は修飾形式として正常動作
{
    const ev = evalVBASingle(`
Function Test() As Long
    Dim d As Object
    Set d = New Scripting.Dictionary
    d.Add "key", 42
    Test = d("key")
End Function`);
    const result = ev.callProcedure('Test', []);
    assert.strictEqual(result, 42, 'New Scripting.Dictionary → 42');
    console.log('[PASS] New Scripting.Dictionary → 正常 (42)');
}

// ---------------------------------------------------------------------------
// registerExternalObject でカスタムプロジェクト名を登録するとエラー検出される
// ---------------------------------------------------------------------------
{
    const ev = new Evaluator(console.log);
    ev.registerComObject((): VbaComObject & { getValue: () => number } => ({ __progId__: 'MyLib.MyClass', getValue: () => 99 }));
    const ast = new Parser(new Lexer(`
Function Test() As Long
    Test = VarType(MyLib)
End Function`).tokenize()).parse();
    ev.evaluateModule(ast);
    ev.resolveIdentifiers([{ ast, moduleName: '' }]);
    try {
        ev.callProcedure('Test', []);
        assert.fail('VarType(MyLib) should throw');
    } catch (e: any) {
        assert.ok(e.message.includes('project'), `MyLib はプロジェクトとしてエラー: "${e.message}"`);
        console.log('[PASS] registerComObject("MyLib.MyClass") → VarType(MyLib) がエラー');
    }
}
{
    const ev = new Evaluator(console.log);
    ev.registerComObject((): VbaComObject & { getValue: () => number } => ({ __progId__: 'MyLib.MyClass', getValue: () => 99 }));
    const ast = new Parser(new Lexer(`
Function Test() As Long
    Dim obj As Object
    Set obj = New MyLib.MyClass
    Test = obj.getValue()
End Function`).tokenize()).parse();
    ev.evaluateModule(ast);
    ev.resolveIdentifiers([{ ast, moduleName: '' }]);
    const result = ev.callProcedure('Test', []);
    assert.strictEqual(result, 99, 'New MyLib.MyClass → 99');
    console.log('[PASS] New MyLib.MyClass → 正常 (99)');
}

console.log('\n✅ namespace-as-value-error: 全テスト通過');
