import { Lexer } from '../../src/compiler/lexer';
import { Parser } from '../../src/compiler/parser';
import { Evaluator } from '../../src/compiler/evaluator';
import { assert } from '../ts/test-runner';

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
    assert.isTrue(evalExpr('CBool(1)'), 'CBool(1)');
    assert.isTrue(evalExpr('1 = 1'), 'Equality 1 = 1');
    assert.isFalse(evalExpr('CBool(0)'), 'CBool(0)');
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
    assert.isTrue(evalExpr('IsNull(nothing)'), 'IsNull(Nothing)');
    assert.isFalse(evalExpr('IsNull(123)'), 'IsNull(123)');
    assert.isTrue(evalExpr('IsArray(Array(1, 2))'), 'IsArray(Array())');
    assert.isFalse(evalExpr('IsArray(123)'), 'IsArray(123)');
    assert.isTrue(evalExpr('IsObject(CreateObject("Scripting.Dictionary"))'), 'IsObject(Dictionary)');
    assert.isFalse(evalExpr('IsObject(123)'), 'IsObject(123)');
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

console.log('\n✅ Built-in Functions: 全テスト通過');
