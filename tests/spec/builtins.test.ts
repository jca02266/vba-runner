import { Lexer } from '../../src/engine/lexer';
import { Parser } from '../../src/engine/parser';
import { Evaluator, vbaTrue, vbaFalse } from '../../src/engine/evaluator';
import { assert } from '../../test-libs/test-runner';

function evalExpr(expr: string): any {
    const tokens = new Lexer(expr).tokenize();
    const parser = new Parser(tokens);
    const ast = (parser as any).parseExpression();
    const ev = new Evaluator(() => {});
    return (ev as any).evaluateExpression(ast);
}

// 1. String Functions
{
    assert.strictEqual(evalExpr('Len("Hello")'), 5, 'Len');
    assert.strictEqual(evalExpr('Left("Hello", 2)'), 'He', 'Left');
    assert.strictEqual(evalExpr('Right("Hello", 2)'), 'lo', 'Right');
    assert.strictEqual(evalExpr('Mid("Hello", 2, 2)'), 'el', 'Mid(start, len)');
    assert.strictEqual(evalExpr('Mid("Hello", 2)'), 'ello', 'Mid(start)');
    assert.strictEqual(evalExpr('InStr("Hello", "e")'), 2, 'InStr found');
    assert.strictEqual(evalExpr('InStr("Hello", "z")'), 0, 'InStr not found');
    assert.strictEqual(evalExpr('LCase("HELLO")'), 'hello', 'LCase');
    assert.strictEqual(evalExpr('Replace("Hello", "e", "a")'), 'Hallo', 'Replace');
}

// 2. Conversion Functions
{
    assert.strictEqual(evalExpr('CInt("123.5")'), 124, 'CInt (round)');
    assert.strictEqual(evalExpr('CStr(123)'), '123', 'CStr');
    assert.strictEqual(evalExpr('CBool(1)'), vbaTrue, 'CBool(1)');
    assert.strictEqual(evalExpr('1 = 1'), vbaTrue, 'Equality 1 = 1');
    assert.strictEqual(evalExpr('CBool(0)'), vbaFalse, 'CBool(0)');
    assert.strictEqual(evalExpr('Fix(123.9)'), 123, 'Fix positive');
    assert.strictEqual(evalExpr('Fix(-123.9)'), -123, 'Fix negative');
}

// 3. Math Functions
{
    assert.strictEqual(evalExpr('Abs(-10)'), 10, 'Abs');
    assert.strictEqual(evalExpr('Round(123.456, 2)'), 123.46, 'Round(n, m)');
    assert.strictEqual(evalExpr('Sqr(16)'), 4, 'Sqr');
}

// 4. Information Functions
{
    assert.strictEqual(evalExpr('IsNull(Null)'), vbaTrue, 'IsNull(Null)');
    assert.strictEqual(evalExpr('IsNull(123)'), vbaFalse, 'IsNull(123)');
    assert.strictEqual(evalExpr('IsArray(Array(1, 2))'), vbaTrue, 'IsArray(Array())');
    assert.strictEqual(evalExpr('IsArray(123)'), vbaFalse, 'IsArray(123)');
    assert.strictEqual(evalExpr('IsObject(CreateObject("Scripting.Dictionary"))'), vbaTrue, 'IsObject(Dictionary)');
    assert.strictEqual(evalExpr('IsObject(123)'), vbaFalse, 'IsObject(123)');
}

// 5. Array Functions
{
    assert.strictEqual(evalExpr('LBound(Array(1, 2, 3))'), 0, 'LBound');
    assert.strictEqual(evalExpr('Array(1, 2, 3)(0)'), 1, 'Array index 0');
    assert.strictEqual(evalExpr('Array(1, 2, 3)(2)'), 3, 'Array index 2');
}

// 6. IIf
{
    assert.strictEqual(evalExpr('IIf(True, "A", "B")'), 'A', 'IIf True');
    assert.strictEqual(evalExpr('IIf(False, "A", "B")'), 'B', 'IIf False');
}

// RGB / QBColor / Nz
{
    assert.strictEqual(evalExpr('RGB(255, 0, 0)'), 255, 'RGB red');
    assert.strictEqual(evalExpr('RGB(0, 255, 0)'), 65280, 'RGB green');
    assert.strictEqual(evalExpr('RGB(0, 0, 255)'), 16711680, 'RGB blue');
    assert.strictEqual(evalExpr('RGB(0, 0, 0)'), 0, 'RGB black');
    assert.strictEqual(evalExpr('RGB(255, 255, 255)'), 16777215, 'RGB white');
    assert.strictEqual(evalExpr('QBColor(0)'), 0, 'QBColor(0) = black');
    assert.strictEqual(evalExpr('QBColor(15)'), 16777215, 'QBColor(15) = white');
    assert.strictEqual(evalExpr('QBColor(4)'), 128, 'QBColor(4) = dark red');
    assert.strictEqual(evalExpr('Nz(Null, 99)'), 99, 'Nz(Null, 99) = 99');
    assert.strictEqual(evalExpr('Nz(42, 99)'), 42, 'Nz(42, 99) = 42');
    assert.strictEqual(evalExpr('Nz(Null)'), 0, 'Nz(Null) default = 0');
    console.log('[PASS] RGB / QBColor / Nz');
}

console.log('\n✅ Built-in Functions: 全テスト通過');
