import { Lexer } from '../../src/engine/lexer';
import { Parser } from '../../src/engine/parser';
import { astEqual, serializeAst, findMatchingExpressions } from '../../src/lsp/ast-comparison';

function parseExpression(exprText: string) {
    const tokens = new Lexer(exprText).tokenize();
    const parser = new Parser(tokens);
    return parser.parseExpressionPublic();
}

function parseCode(code: string) {
    const tokens = new Lexer(code).tokenize();
    const parser = new Parser(tokens);
    return parser.parse();
}

let passed = 0;
let failed = 0;

function test(name: string, condition: boolean, message?: string) {
    if (condition) {
        console.log(`✓ ${name}`);
        passed++;
    } else {
        console.log(`✗ ${name}: ${message || 'assertion failed'}`);
        failed++;
    }
}

console.log('Testing AST Comparison\n');

// Test 1: Identical expressions
const expr1a = parseExpression('a + b');
const expr1b = parseExpression('a + b');
test('1. Identical expressions (a + b)', astEqual(expr1a, expr1b));

// Test 2: Expressions with different spacing
const expr2a = parseExpression('a+b');
const expr2b = parseExpression('a + b');
test('2. Expressions with different spacing (a+b vs a + b)', astEqual(expr2a, expr2b));

// Test 3: Different operators
const expr3a = parseExpression('a + b');
const expr3b = parseExpression('a - b');
test('3. Different operators (a + b vs a - b)', !astEqual(expr3a, expr3b));

// Test 4: Different identifiers
const expr4a = parseExpression('a + b');
const expr4b = parseExpression('c + d');
test('4. Different identifiers (a + b vs c + d)', !astEqual(expr4a, expr4b));

// Test 5: Function calls (same)
const expr5a = parseExpression('CalcSum(a, b)');
const expr5b = parseExpression('CalcSum(a, b)');
test('5. Function calls - same (CalcSum(a, b))', astEqual(expr5a, expr5b));

// Test 6: Function calls (different arguments)
const expr6a = parseExpression('CalcSum(a, b)');
const expr6b = parseExpression('CalcSum(a, c)');
test('6. Function calls - different arguments', !astEqual(expr6a, expr6b));

// Test 7: Complex expressions (identical)
const expr7a = parseExpression('CalcSum(a, b) + CalcSum(c, d)');
const expr7b = parseExpression('CalcSum(a, b) + CalcSum(c, d)');
test('7. Complex expressions - identical', astEqual(expr7a, expr7b));

// Test 8: Complex expressions (different order)
const expr8a = parseExpression('a + b * c');
const expr8b = parseExpression('b * c + a');
test('8. Complex expressions - different order', !astEqual(expr8a, expr8b));

// Test 9: Member access (same)
const expr9a = parseExpression('obj.prop');
const expr9b = parseExpression('obj.prop');
test('9. Member access - same (obj.prop)', astEqual(expr9a, expr9b));

// Test 10: Member access (different)
const expr10a = parseExpression('obj.prop1');
const expr10b = parseExpression('obj.prop2');
test('10. Member access - different', !astEqual(expr10a, expr10b));

// Test 11: Finding matching expressions in code
const code11 = `Sub Test()
    Dim a, b, total
    a = 5
    b = 10
    total = CalcSum(a, b) + CalcSum(a, b)
    Call Other(CalcSum(a, b))
End Sub`;
const ast11 = parseCode(code11);
const targetExpr11 = parseExpression('CalcSum(a, b)');
const proc11 = ast11.body.find((s: any) => s.type === 'ProcedureDeclaration') as any;
const matches11 = findMatchingExpressions(proc11?.body || [], targetExpr11);
test('11. Finding matching expressions in procedure', matches11.length >= 2, `found ${matches11.length} matches (expected >=2)`);

// Test 12: Not matching different arguments
const code12 = `Sub Test()
    total = CalcSum(a, b) + CalcSum(c, d)
End Sub`;
const ast12 = parseCode(code12);
const targetExpr12 = parseExpression('CalcSum(a, b)');
const proc12 = ast12.body.find((s: any) => s.type === 'ProcedureDeclaration') as any;
const matches12 = findMatchingExpressions(proc12?.body || [], targetExpr12);
test('12. Different function arguments not matched', matches12.length === 1, `found ${matches12.length} matches (expected 1)`);

console.log(`\n${passed} passed, ${failed} failed`);
if (failed > 0) {
    process.exit(1);
}
