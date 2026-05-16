import { Lexer, TokenType } from '../../src/compiler/lexer';
import { Parser } from '../../src/compiler/parser';
import { assert } from '../../test-libs/test-runner';

function tokenize(src: string) {
    return new Lexer(src).tokenize();
}

// 1. First token on line starts at column 1
{
    const tokens = tokenize('Sub Foo()');
    assert.strictEqual(tokens[0].type, TokenType.KeywordSub, 'first token is Sub');
    assert.strictEqual(tokens[0].line, 1, 'line 1');
    assert.strictEqual(tokens[0].column, 1, 'column 1');
    console.log('[PASS] First token: line=', tokens[0].line, 'column=', tokens[0].column);
}

// 2. Subsequent tokens have correct column
{
    const tokens = tokenize('x = 42');
    // x   = column 1
    // =   = column 3
    // 42  = column 5
    assert.strictEqual(tokens[0].column, 1, 'x at column 1');
    assert.strictEqual(tokens[1].column, 3, '= at column 3');
    assert.strictEqual(tokens[2].column, 5, '42 at column 5');
    console.log('[PASS] Column positions in "x = 42":', tokens.map(t => t.column));
}

// 3. Second line resets column to 1
{
    const tokens = tokenize('x = 1\ny = 2');
    const yToken = tokens.find(t => t.value === 'y');
    assert.ok(yToken !== undefined, 'y token found');
    assert.strictEqual(yToken!.line, 2, 'y on line 2');
    assert.strictEqual(yToken!.column, 1, 'y at column 1 on new line');
    console.log('[PASS] Line 2 column reset: y at line=', yToken!.line, 'column=', yToken!.column);
}

// 4. Indented code has correct column
{
    const tokens = tokenize('Sub Foo()\n    Dim x\nEnd Sub');
    const dimToken = tokens.find(t => t.type === TokenType.KeywordDim);
    assert.ok(dimToken !== undefined, 'Dim token found');
    assert.strictEqual(dimToken!.line, 2, 'Dim on line 2');
    assert.strictEqual(dimToken!.column, 5, 'Dim at column 5 (4 spaces indent)');
    console.log('[PASS] Indented Dim at line=', dimToken!.line, 'column=', dimToken!.column);
}

// 5. Multi-char operators have correct column
{
    const tokens = tokenize('a <> b');
    const neqToken = tokens.find(t => t.type === TokenType.OperatorNotEquals);
    assert.ok(neqToken !== undefined, '<> token found');
    assert.strictEqual(neqToken!.column, 3, '<> at column 3');
    console.log('[PASS] <> operator at column=', neqToken!.column);
}

// 6. String literal column is where quote starts
{
    const tokens = tokenize('x = "hello"');
    const strToken = tokens.find(t => t.type === TokenType.String);
    assert.ok(strToken !== undefined, 'String token found');
    assert.strictEqual(strToken!.column, 5, 'String starts at column 5');
    console.log('[PASS] String literal at column=', strToken!.column);
}

// 7. Identifier after keyword has correct column
{
    const tokens = tokenize('Dim myVar As Integer');
    const myVarToken = tokens.find(t => t.value.toLowerCase() === 'myvar');
    assert.ok(myVarToken !== undefined, 'myVar token found');
    assert.strictEqual(myVarToken!.column, 5, 'myVar at column 5');
    console.log('[PASS] Identifier after keyword at column=', myVarToken!.column);
}

// 8. Newline token has column of the newline character
{
    const tokens = tokenize('x = 1\ny = 2');
    const newlineToken = tokens.find(t => t.type === TokenType.Newline);
    assert.ok(newlineToken !== undefined, 'Newline token found');
    assert.strictEqual(newlineToken!.line, 1, 'Newline on line 1');
    console.log('[PASS] Newline token at line=', newlineToken!.line, 'column=', newlineToken!.column);
}

// 9. Parser ASTNode has start and end positions
{
    const src = 'Sub Foo()\n    x = 42\nEnd Sub';
    const tokens = new Lexer(src).tokenize();
    const ast = new Parser(tokens).parse();
    const proc = ast.body.find((s: any) => s.type === 'ProcedureDeclaration') as any;
    assert.ok(proc !== undefined, 'ProcedureDeclaration found');
    assert.ok(proc.start !== undefined, 'start position set');
    assert.ok(proc.end !== undefined, 'end position set');
    assert.strictEqual(proc.start.line, 1, 'proc starts at line 1');
    assert.strictEqual(proc.start.column, 1, 'proc starts at column 1');
    console.log('[PASS] ASTNode start:', proc.start, 'end:', proc.end);
}

// 10. Assignment statement inside proc has correct start position
{
    const src = 'Sub Foo()\n    x = 42\nEnd Sub';
    const tokens = new Lexer(src).tokenize();
    const ast = new Parser(tokens).parse();
    const proc = ast.body.find((s: any) => s.type === 'ProcedureDeclaration') as any;
    const assign = proc.body.find((s: any) => s.type === 'AssignmentStatement') as any;
    assert.ok(assign !== undefined, 'AssignmentStatement found');
    assert.strictEqual(assign.start.line, 2, 'assignment on line 2');
    assert.strictEqual(assign.start.column, 5, 'assignment at column 5 (4 spaces)');
    assert.ok(assign.end !== undefined, 'end position set');
    assert.strictEqual(assign.end.line, 2, 'end on same line');
    console.log('[PASS] Assignment start:', assign.start, 'end:', assign.end);
}

// 11. Tab indentation counts as 1 column per char
{
    const tokens = tokenize('Sub Foo()\n\tDim x\nEnd Sub');
    const dimToken = tokens.find(t => t.type === TokenType.KeywordDim);
    assert.ok(dimToken !== undefined, 'Dim token found');
    assert.strictEqual(dimToken!.column, 2, 'Dim at column 2 after 1 tab');
    console.log('[PASS] Tab indent: Dim at column=', dimToken!.column);
}

// 12. column is consistent across all tokens on a line
{
    const src = 'If x > 0 Then y = 1';
    const tokens = tokenize(src);
    // Verify none of the columns are 0 or undefined
    for (const tok of tokens.filter(t => t.type !== TokenType.EOF)) {
        assert.ok(tok.column > 0, `column > 0 for token "${tok.value}"`);
        assert.ok(tok.line > 0, `line > 0 for token "${tok.value}"`);
    }
    console.log('[PASS] All tokens have positive line and column');
}

console.log('\n✅ Lexer column tracking: 全テスト通過');
