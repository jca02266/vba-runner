import { Lexer } from '../../src/engine/lexer';
import { Parser } from '../../src/engine/parser';
import { Evaluator } from '../../src/engine/evaluator';
import { assert } from '../../test-libs/test-runner';

function evalExpr(expr: string): any {
    const ev = new Evaluator(console.log);
    return ev.evalExpression(expr);
}

console.log("--- Starting Int & Val Function Tests ---");

// 1. Int
assert.strictEqual(evalExpr('Int(1.1)'), 1, 'Int(1.1)');
assert.strictEqual(evalExpr('Int(1.9)'), 1, 'Int(1.9)');
assert.strictEqual(evalExpr('Int(-1.1)'), -2, 'Int(-1.1) should floor');
assert.strictEqual(evalExpr('Int(-1.9)'), -2, 'Int(-1.9)');

// 2. Val
assert.strictEqual(evalExpr('Val("123")'), 123, 'Val basic');
assert.strictEqual(evalExpr('Val("  123.45 abc")'), 123.45, 'Val with spaces and trailing chars');
assert.strictEqual(evalExpr('Val("&H10")'), 16, 'Val Hex');
assert.strictEqual(evalExpr('Val("&O10")'), 8, 'Val Octal');
assert.strictEqual(evalExpr('Val("abc")'), 0, 'Val no numeric');
assert.strictEqual(evalExpr('Val("1.5e2")'), 150, 'Val scientific notation 1.5e2');
assert.strictEqual(evalExpr('Val("2.0E+3")'), 2000, 'Val scientific notation 2.0E+3');
assert.strictEqual(evalExpr('Val("3.5E-1")'), 0.35, 'Val scientific notation 3.5E-1');

// Val: 非文字列引数は String に強制変換してから解析 (VBA 仕様)
assert.strictEqual(evalExpr('Val(42)'), 42, 'Val(42) coerces number to "42"');
assert.strictEqual(evalExpr('Val(3.14)'), 3.14, 'Val(3.14) coerces to "3.14"');
// Val(True): CStr(True)="True" → Val("True")=0 (VBA does NOT convert Boolean to "-1"/0 for Val)
assert.strictEqual(evalExpr('Val(True)'), 0, 'Val(True) coerces to "True" → 0');
assert.strictEqual(evalExpr('Val(False)'), 0, 'Val(False) coerces to "False" → 0');

// Bug CX: Hex 負数の桁数 (§6.1.2.3.1.17)
// 仕様: -32767〜-1 → 4文字（16ビット2の補数）、-2147483648〜-32768 → 8文字（32ビット）
{
    assert.strictEqual(evalExpr('Hex(-1)'), 'FFFF', 'Hex(-1) → FFFF (16ビット4文字)');
    assert.strictEqual(evalExpr('Hex(-32767)'), '8001', 'Hex(-32767) → 8001 (16ビット4文字)');
    assert.strictEqual(evalExpr('Hex(-32768)'), 'FFFF8000', 'Hex(-32768) → FFFF8000 (32ビット8文字)');
    assert.strictEqual(evalExpr('Hex(-2147483648)'), '80000000', 'Hex(-2147483648) → 80000000 (32ビット8文字)');
    assert.strictEqual(evalExpr('Hex(255)'), 'FF', 'Hex(255) 正数は先頭ゼロなし');
    console.log('[PASS] Bug CX: Hex 負数 16ビット範囲は4文字');
}

// Bug CY: Oct 負数の桁数 (§6.1.2.3.1.19)
// 仕様: -32767〜-1 → 6文字（16ビット2の補数8進）、-2147483648〜-32768 → 11文字
{
    assert.strictEqual(evalExpr('Oct(-1)'), '177777', 'Oct(-1) → 177777 (16ビット6文字)');
    assert.strictEqual(evalExpr('Oct(-32767)'), '100001', 'Oct(-32767) → 100001 (16ビット6文字)');
    assert.strictEqual(evalExpr('Oct(-32768)'), '37777700000', 'Oct(-32768) → 37777700000 (32ビット11文字)');
    assert.strictEqual(evalExpr('Oct(8)'), '10', 'Oct(8) 正数は先頭ゼロなし');
    console.log('[PASS] Bug CY: Oct 負数 16ビット範囲は6文字');
}

console.log("✅ Int & Val: All tests passed!");
