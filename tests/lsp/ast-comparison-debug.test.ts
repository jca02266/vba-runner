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

// Debug Test 2: Spacing
console.log('=== Debug Test 2: Spacing ===');
const expr2a = parseExpression('a+b');
const expr2b = parseExpression('a + b');
const ser2a = serializeAst(expr2a);
const ser2b = serializeAst(expr2b);
console.log('a+b serialized:', ser2a);
console.log('a + b serialized:', ser2b);
console.log('Equal?', ser2a === ser2b);
console.log('astEqual?', astEqual(expr2a, expr2b));

// Debug Test 11: Find expressions
console.log('\n=== Debug Test 11: Find Expressions ===');
const code11 = `Sub Test()
    total = CalcSum(a, b) + CalcSum(a, b)
End Sub`;
const ast11 = parseCode(code11);
console.log('AST body:', ast11.body.length);
const proc11 = ast11.body.find((s: any) => s.type === 'ProcedureDeclaration') as any;
console.log('Procedure found:', proc11 ? 'yes' : 'no');
console.log('Procedure body:', proc11?.body ? proc11.body.length : 0);

if (proc11?.body) {
    console.log('\nWalking expressions...');
    let exprCount = 0;
    const walk = (node: any, depth = 0) => {
        if (!node || typeof node !== 'object') return;
        const isExpression = node.type && (
            node.type.includes('Expression') ||
            node.type === 'BinaryOp' ||
            node.type === 'Identifier' ||
            node.type === 'Literal' ||
            node.type === 'CallExpression' ||
            node.type === 'MemberExpression'
        );
        if (isExpression) {
            exprCount++;
            console.log(`  ${depth}: ${node.type} at line ${node.loc?.start.line}`);
        }
        if (Array.isArray(node)) {
            for (const item of node) walk(item, depth + 1);
        } else {
            for (const key of Object.keys(node)) {
                if (key !== 'loc') {
                    walk(node[key], depth + 1);
                }
            }
        }
    };
    walk(proc11.body);
    console.log(`Total expressions found: ${exprCount}`);
}

const targetExpr11 = parseExpression('CalcSum(a, b)');
const matches11 = findMatchingExpressions(proc11?.body || [], targetExpr11);
console.log(`\nMatches found: ${matches11.length}`);
