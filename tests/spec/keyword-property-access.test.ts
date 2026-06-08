import { Lexer } from '../../src/engine/lexer';
import { Parser } from '../../src/engine/parser';
import { evalVBASingle, assert } from '../../test-libs/test-runner';

// Bug: `With obj: .Collection = x: End With` failed because the implicit-With
// member access handler checked `propToken.type !== Identifier`. When the property
// name was a VBA keyword token (KeywordCollection etc.), the check threw an error
// and error recovery aborted the containing function.

function evalVBA(code: string): any {
    return evalVBASingle(code);
}

function parseStatements(src: string) {
    const tokens = new Lexer(src).tokenize();
    return new Parser(tokens).parse().body;
}

// Test 1: .Collection as implicit-With property (parse only)
{
    const stmts = parseStatements(`
Sub S()
    With obj
        .Collection = 1
    End With
End Sub
`);
    assert.strictEqual(stmts.length, 1, '.Collection in With block does not abort Sub');
    assert.strictEqual(stmts[0].type, 'ProcedureDeclaration', 'ProcedureDeclaration');
    console.log('[PASS] .Collection in With block: parses without error');
}

// Test 2: .keyword implicit-With — no statements leak outside the Sub
{
    const src = `
Sub S()
    With obj
        .Collection = 1
        .Name = "test"
        .Type = 2
    End With
End Sub
Sub Other()
End Sub
`;
    const tokens = new Lexer(src).tokenize();
    const ast = new Parser(tokens).parse();
    assert.strictEqual(ast.body.length, 2, 'Both Subs present — nothing leaked to top level');
    console.log('[PASS] .keyword in With: no statements leak to top level');
}

// Test 3: Keyword property names on regular member access (not implicit-With)
//   obj.Item(...) where Item is not a keyword, but ensure keyword-named properties
//   work the same way via the postfix loop (which already accepted keywords)
{
    const stmts = parseStatements(`
Sub S()
    x = obj.String
    y = obj.Integer
End Sub
`);
    assert.strictEqual(stmts.length, 1, 'obj.String / obj.Integer parse without error');
    console.log('[PASS] obj.keyword-property: member access with keyword property name');
}

// Test 4: .Name (Name is a keyword) in implicit With
{
    const stmts = parseStatements(`
Sub S()
    With obj
        .Name = "test"
    End With
End Sub
`);
    assert.strictEqual(stmts.length, 1, '.Name (keyword) in With does not abort Sub');
    console.log('[PASS] .Name (keyword) in With block: parses');
}

console.log('\n✅ keyword-property-access: 全テスト通過');
