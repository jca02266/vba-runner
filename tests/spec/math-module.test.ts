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

function runFunc(code: string, name: string, args: any[] = []): any {
    return evalVBA(code).callProcedure(name, args);
}

const PI = Math.PI;

// --- Sgn ---
{
    const code = `
Function T(n As Double) As Integer: T = Sgn(n): End Function
`;
    assert.strictEqual(runFunc(code, 'T', [5]),    1,  'Sgn(正) = 1');
    assert.strictEqual(runFunc(code, 'T', [0]),    0,  'Sgn(0) = 0');
    assert.strictEqual(runFunc(code, 'T', [-3]),  -1,  'Sgn(負) = -1');
    assert.strictEqual(runFunc(code, 'T', [0.001]), 1, 'Sgn(小正) = 1');
    console.log('[PASS] Sgn');
}

// --- Atn ---
{
    const code = `
Function T(n As Double) As Double: T = Atn(n): End Function
`;
    const r = runFunc(code, 'T', [1]) as number;
    assert.strictEqual(Math.abs(r - PI / 4) < 1e-10, true, 'Atn(1) ≈ π/4');

    const r2 = runFunc(code, 'T', [0]) as number;
    assert.strictEqual(r2, 0, 'Atn(0) = 0');

    const rNeg = runFunc(code, 'T', [-1]) as number;
    assert.strictEqual(Math.abs(rNeg - (-PI / 4)) < 1e-10, true, 'Atn(-1) ≈ -π/4');
    console.log('[PASS] Atn');
}

// --- Cos ---
{
    const code = `
Function T(n As Double) As Double: T = Cos(n): End Function
`;
    assert.strictEqual(Math.abs((runFunc(code, 'T', [0]) as number) - 1) < 1e-10,     true, 'Cos(0) = 1');
    assert.strictEqual(Math.abs((runFunc(code, 'T', [PI]) as number) - (-1)) < 1e-10, true, 'Cos(π) = -1');
    assert.strictEqual(Math.abs(runFunc(code, 'T', [PI / 2]) as number) < 1e-10,      true, 'Cos(π/2) ≈ 0');
    console.log('[PASS] Cos');
}

// --- Sin ---
{
    const code = `
Function T(n As Double) As Double: T = Sin(n): End Function
`;
    assert.strictEqual(Math.abs((runFunc(code, 'T', [0]) as number)) < 1e-10,             true, 'Sin(0) = 0');
    assert.strictEqual(Math.abs((runFunc(code, 'T', [PI / 2]) as number) - 1) < 1e-10,   true, 'Sin(π/2) = 1');
    assert.strictEqual(Math.abs((runFunc(code, 'T', [PI]) as number)) < 1e-10,            true, 'Sin(π) ≈ 0');
    assert.strictEqual(Math.abs((runFunc(code, 'T', [-PI / 2]) as number) + 1) < 1e-10,  true, 'Sin(-π/2) = -1');
    console.log('[PASS] Sin');
}

// --- Tan ---
{
    const code = `
Function T(n As Double) As Double: T = Tan(n): End Function
`;
    assert.strictEqual(Math.abs((runFunc(code, 'T', [0]) as number)) < 1e-10,             true, 'Tan(0) = 0');
    assert.strictEqual(Math.abs((runFunc(code, 'T', [PI / 4]) as number) - 1) < 1e-10,   true, 'Tan(π/4) = 1');
    assert.strictEqual(Math.abs((runFunc(code, 'T', [-PI / 4]) as number) + 1) < 1e-10,  true, 'Tan(-π/4) = -1');
    console.log('[PASS] Tan');
}

// --- Exp ---
{
    const code = `
Function T(n As Double) As Double: T = Exp(n): End Function
`;
    assert.strictEqual(Math.abs((runFunc(code, 'T', [0]) as number) - 1) < 1e-10,         true, 'Exp(0) = 1');
    assert.strictEqual(Math.abs((runFunc(code, 'T', [1]) as number) - Math.E) < 1e-10,    true, 'Exp(1) = e');
    assert.strictEqual(Math.abs((runFunc(code, 'T', [-1]) as number) - 1 / Math.E) < 1e-10, true, 'Exp(-1) = 1/e');
    assert.strictEqual((runFunc(code, 'T', [10]) as number) > 22000, true, 'Exp(10) > 22000');
    console.log('[PASS] Exp');
}

// --- Log ---
{
    const code = `
Function T(n As Double) As Double: T = Log(n): End Function
`;
    assert.strictEqual(Math.abs((runFunc(code, 'T', [1]) as number)) < 1e-10,             true, 'Log(1) = 0');
    assert.strictEqual(Math.abs((runFunc(code, 'T', [Math.E]) as number) - 1) < 1e-10,   true, 'Log(e) = 1');
    assert.strictEqual((runFunc(code, 'T', [100]) as number) > 4.6, true, 'Log(100) > 4.6');
    // Log(0) や Log(負) はエラーだが仕様では未定義動作なのでここでは確認のみ
    console.log('[PASS] Log');
}

// --- Rnd ---
{
    // Rnd() は [0, 1) の範囲に収まる
    const code = `
Function GetRnd() As Double: GetRnd = Rnd(): End Function
Function GetRndArg(n As Double) As Double: GetRndArg = Rnd(n): End Function
Function GetRndZero() As Double: GetRndZero = Rnd(0): End Function
`;
    const ev = evalVBA(code);
    const v1 = ev.callProcedure('GetRnd', []) as number;
    assert.strictEqual(v1 >= 0 && v1 < 1, true, 'Rnd() ∈ [0, 1)');

    const v2 = ev.callProcedure('GetRnd', []) as number;
    // 連続呼び出しで値が得られる（同じになる確率はほぼ0）
    assert.strictEqual(typeof v2 === 'number', true, 'Rnd() は数値');

    // Rnd(0) は最後に生成された値を返す（同じ値）
    const last = ev.callProcedure('GetRnd', []) as number;
    const fromZero = ev.callProcedure('GetRndZero', []) as number;
    assert.strictEqual(last, fromZero, 'Rnd(0) は最後の値を返す');

    // Rnd(負) は固定シードから決定論的な値を返す
    const neg1a = ev.callProcedure('GetRndArg', [-1]) as number;
    const neg1b = ev.callProcedure('GetRndArg', [-1]) as number;
    assert.strictEqual(neg1a, neg1b, 'Rnd(-1) は同じ値を繰り返す');
    assert.strictEqual(neg1a >= 0 && neg1a < 1, true, 'Rnd(負) も [0, 1) 範囲');
    console.log('[PASS] Rnd');
}

// --- Randomize / Rnd の組み合わせ ---
{
    const code = `
Function GetRndAfterRandomize() As Double
    Randomize 42
    GetRndAfterRandomize = Rnd()
End Function
`;
    const r1 = runFunc(code, 'GetRndAfterRandomize') as number;
    assert.strictEqual(r1 >= 0 && r1 < 1, true, 'Randomize後のRnd() ∈ [0, 1)');
    // 同じシードで Randomize すれば同じ値が得られる
    const r2 = runFunc(code, 'GetRndAfterRandomize') as number;
    assert.strictEqual(r1, r2, '同一シードで同じ値');
    console.log('[PASS] Randomize');
}

// --- 三角関数・逆三角関数の組み合わせ確認 ---
{
    const code = `
Function TestPi() As Double
    Dim pi As Double
    pi = 4 * Atn(1)  ' VBA慣用の π の計算法
    TestPi = pi
End Function

Function TestSinCos() As Double
    Dim pi As Double
    pi = 4 * Atn(1)
    TestSinCos = Sin(pi / 6) + Cos(pi / 3)
End Function
`;
    const piVal = runFunc(code, 'TestPi') as number;
    assert.strictEqual(Math.abs(piVal - PI) < 1e-10, true, '4*Atn(1) = π');

    // Sin(30°) + Cos(60°) = 0.5 + 0.5 = 1.0
    const sc = runFunc(code, 'TestSinCos') as number;
    assert.strictEqual(Math.abs(sc - 1.0) < 1e-10, true, 'Sin(π/6)+Cos(π/3) = 1');
    console.log('[PASS] 三角関数組み合わせ (4*Atn(1)=π)');
}

console.log('\n✅ Math Module: 全テスト通過');
