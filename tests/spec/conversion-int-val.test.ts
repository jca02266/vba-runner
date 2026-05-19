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

console.log("✅ Int & Val: All tests passed!");
