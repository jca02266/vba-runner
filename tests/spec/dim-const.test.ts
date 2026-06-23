/**
 * Dim, Const, Let, Set (§5.4.2.1, §5.4.2.2, §5.4.3.8, §5.4.3.2) のテスト
 */
import { evalVBASingle, assert, assertCompileErrorResolve, assertCompileErrorPreproc } from '../../test-libs/test-runner';

// --- 1. Dim と初期値 ---
const dimCode = `
    Public i As Integer
    Public s As String
    Public b As Boolean
    Public v
    
    Sub Test()
    End Sub
`;
const ev1 = evalVBASingle(dimCode);
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
const ev2 = evalVBASingle(constCode);
assert.strictEqual(ev2.env.get('my_pi'), 3.14, 'Const 定数が取得できる');
assert.strictEqual(ev2.env.get('app_name'), "VBA", '型付き Const が取得できる');

const constErrCode = `
    Const X = 10
    Sub Fail()
        X = 20
    End Sub
`;
const ev2e = evalVBASingle(constErrCode);
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
const ev3 = evalVBASingle(letCode);
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
const ev5 = evalVBASingle(setGlobalCode);
ev5.callProcedure('TestSet', []);
const gObj = ev5.env.get('gobj');
assert.strictEqual(gObj.__instanceEnv__.get('value'), "Modified", 'Set による参照代入とプロパティ更新');
console.log('[PASS] Set (オブジェクト代入)');

// --- 5. 配列境界式に Const を参照する ---
// モジュールレベル 1D 配列
{
    const ev = evalVBASingle(`
Const W As Integer = 3
Dim a(0 To W - 1) As Integer
Sub Test()
    a(2) = 5
End Sub`);
    ev.callProcedure('Test', []);
    const arr = ev.env.get('a');
    assert.strictEqual((arr as any).__vbaDimensions__[0].upper, 2, 'モジュールレベル 1D: UBound = W-1 = 2');
    assert.strictEqual(arr[2], 5, 'モジュールレベル 1D: a(2) = 5 が読み書きできる');
    console.log('[PASS] 配列境界に Const を参照: モジュールレベル 1D');
}

// モジュールレベル 2D 配列
{
    const ev = evalVBASingle(`
Const W As Integer = 3
Dim a(0 To W - 1, 0 To 3) As Integer
Sub Test()
    a(2, 3) = 7
End Sub`);
    ev.callProcedure('Test', []);
    const arr = ev.env.get('a');
    assert.strictEqual((arr as any).__vbaDimensions__[0].upper, 2, 'モジュールレベル 2D: dim1 UBound = 2');
    assert.strictEqual((arr as any).__vbaDimensions__[1].upper, 3, 'モジュールレベル 2D: dim2 UBound = 3');
    assert.strictEqual(arr[2][3], 7, 'モジュールレベル 2D: a(2,3) = 7 が読み書きできる');
    console.log('[PASS] 配列境界に Const を参照: モジュールレベル 2D');
}

// Sub 内 Dim で Public Const を参照
{
    const ev = evalVBASingle(`
Public Const W As Integer = 3
Sub Test()
    Dim a(0 To W - 1) As Integer
    a(2) = 9
End Sub`);
    ev.callProcedure('Test', []);
    console.log('[PASS] 配列境界に Const を参照: Sub 内 Dim');
}

// --- 6. Dim の配列サイズに変数を指定するとコンパイルエラー ---
// モジュールレベル: n は Dim 変数 → "Compile error: Constant expression required"
assertCompileErrorResolve(
    `Dim n As Integer\nDim a(n) As Integer\nSub Test()\nEnd Sub`,
    undefined,
    /Compile error: Constant expression required/i,
    'モジュールレベル Dim 配列サイズに変数を使うとコンパイルエラー'
);

// プロシージャレベル: n は Dim 変数 → "Compile error: Constant expression required"
assertCompileErrorPreproc(
    `Sub Test()\n    Dim n As Integer\n    n = 5\n    Dim a(n) As Integer\nEnd Sub`,
    'Test',
    undefined,
    /Compile error: Constant expression required/i,
    'プロシージャ内 Dim 配列サイズに変数を使うとコンパイルエラー'
);

// 2次元配列の lower bound に変数を使ってもエラー
assertCompileErrorPreproc(
    `Sub Test()\n    Dim n As Integer\n    n = 1\n    Dim a(n To 5) As Integer\nEnd Sub`,
    'Test',
    undefined,
    /Compile error: Constant expression required/i,
    'Dim 配列 lower bound に変数を使うとコンパイルエラー'
);

// Const を使う場合は OK（既存の挙動を回帰チェック）
{
    const ev = evalVBASingle(`
        Const MAX As Integer = 3
        Dim a(MAX) As Integer
        Sub Test()
            a(1) = 10
            a(MAX) = 30
        End Sub
    `);
    ev.callProcedure('Test', []);
    const arr = ev.env.get('a');
    assert.strictEqual(arr[1], 10, 'Const を使った Dim 配列: a(1) = 10');
    assert.strictEqual(arr[3], 30, 'Const を使った Dim 配列: a(MAX) = 30');
    console.log('[PASS] Dim 配列サイズに Const を使う（回帰）');
}

// 式（Const を含む計算）も OK
{
    const ev = evalVBASingle(`
        Const LO As Integer = 1
        Const HI As Integer = 5
        Dim a(LO To HI) As Integer
        Sub Test()
            a(LO) = 11
            a(HI) = 55
        End Sub
    `);
    ev.callProcedure('Test', []);
    const arr = ev.env.get('a');
    assert.strictEqual(arr[1], 11, 'Const 式の Dim 配列 a(1)');
    assert.strictEqual(arr[5], 55, 'Const 式の Dim 配列 a(5)');
    console.log('[PASS] Dim 配列サイズに Const 式を使う（回帰）');
}

console.log('\n✅ Dim, Const, Let, Set: 全テスト通過');
