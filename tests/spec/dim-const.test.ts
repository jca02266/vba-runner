/**
 * Dim, Const, Let, Set (§5.4.2.1, §5.4.2.2, §5.4.3.8, §5.4.3.2) のテスト
 */
import { Lexer } from '../../src/compiler/lexer';
import { Parser } from '../../src/compiler/parser';
import { Evaluator } from '../../src/compiler/evaluator';
import { assert } from '../ts/test-runner';

function evalVBA(code: string): any {
    const tokens = new Lexer(code).tokenize();
    const ast = new Parser(tokens).parse();
    const ev = new Evaluator(console.log);
    ev.evaluate(ast);
    return ev;
}

// --- 1. Dim と初期値 ---
const dimCode = `
    Public i As Integer
    Public s As String
    Public b As Boolean
    Public v
    
    Sub Test()
    End Sub
`;
const ev1 = evalVBA(dimCode);
assert.strictEqual(ev1.env.get('i'), 0, 'Integer の初期値は 0');
assert.strictEqual(ev1.env.get('s'), "", 'String の初期値は ""');
assert.strictEqual(ev1.env.get('b'), 0, 'Boolean の初期値は 0 (False)');
assert.strictEqual(ev1.env.get('v'), null, 'Variant の初期値は vbaEmpty (null)');
console.log('[PASS] Dim と初期値');

// --- 2. Const と定数保護 ---
const constCode = `
    Public Const MY_PI = 3.14
    Public Const APP_NAME As String = "VBA"
    
    Sub Test()
    End Sub
`;
const ev2 = evalVBA(constCode);
assert.strictEqual(ev2.env.get('my_pi'), 3.14, 'Const 定数が取得できる');
assert.strictEqual(ev2.env.get('app_name'), "VBA", '型付き Const が取得できる');

const constErrCode = `
    Const X = 10
    Sub Fail()
        X = 20
    End Sub
`;
const ev2e = evalVBA(constErrCode);
let threwConst = false;
try {
    ev2e.callProcedure('Fail', []);
} catch (e: any) {
    threwConst = true;
    assert.strictEqual(e.message.includes('constant'), true, '定数への代入でエラーが発生');
}
assert.strictEqual(threwConst, true, '定数保護が機能している');
console.log('[PASS] Const と定数保護');

// --- 3. Let (明示的・暗黙的) ---
const letCode = `
    Public x
    Public y
    Sub Test()
        Let x = 100
        y = 200
    End Sub
`;
const ev3 = evalVBA(letCode);
ev3.callProcedure('Test', []);
assert.strictEqual(ev3.env.get('x'), 100, '明示的 Let');
assert.strictEqual(ev3.env.get('y'), 200, '暗黙的 Let');
console.log('[PASS] Let (明示的・暗黙的)');

// --- 4. Set (オブジェクト代入) ---
const setGlobalCode = `
    Class MyObj
        Public Value
    End Class
    Public gObj
    Sub TestSet()
        Dim o
        Set o = New MyObj
        o.Value = "Original"
        Set gObj = o
        gObj.Value = "Modified"
    End Sub
`;
const ev5 = evalVBA(setGlobalCode);
ev5.callProcedure('TestSet', []);
const gObj = ev5.env.get('gobj');
assert.strictEqual(gObj.__instanceEnv__.get('value'), "Modified", 'Set による参照代入とプロパティ更新');
console.log('[PASS] Set (オブジェクト代入)');

console.log('\n✅ Dim, Const, Let, Set: 全テスト通過');
