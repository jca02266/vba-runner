import { Lexer } from '../../src/compiler/lexer';
import { Parser } from '../../src/compiler/parser';
import { assert } from '../../test-libs/test-runner';

function extractSymbols(src: string): any[] {
    const tokens = new Lexer(src).tokenize();
    const ast = new Parser(tokens).parse();
    return collectSymbols(ast.body);
}

function collectSymbols(statements: any[]): any[] {
    const symbols: any[] = [];
    for (const stmt of statements) {
        if (stmt.type === 'ProcedureDeclaration') {
            symbols.push({
                name: stmt.name.name,
                kind: stmt.isProperty ? 6 : (stmt.isFunction ? 12 : 11), // Property=6, Function=12, Subroutine=11
                range: stmt.loc ? { start: stmt.loc.start, end: stmt.loc.end } : undefined,
                isFunction: stmt.isFunction,
                isProperty: stmt.isProperty
            });
        } else if (stmt.type === 'ClassDeclaration') {
            const members = collectSymbols(stmt.body);
            symbols.push({
                name: stmt.name,
                kind: 5, // Class
                range: undefined,
                children: members
            });
        } else if (stmt.type === 'VariableDeclaration') {
            for (const decl of stmt.declarations) {
                symbols.push({
                    name: decl.name.name,
                    kind: 13, // Variable
                    range: undefined
                });
            }
        }
    }
    return symbols;
}

// 1. Simple Sub procedure is recognized
{
    const syms = extractSymbols('Sub Foo()\nEnd Sub');
    assert.strictEqual(syms.length, 1, 'one symbol');
    assert.strictEqual(syms[0].name, 'Foo', 'name is Foo');
    assert.strictEqual(syms[0].kind, 11, 'kind 11 = Subroutine');
    assert.ok(syms[0].range, 'range present');
    console.log('[PASS] Simple Sub recognized');
}

// 2. Function procedure is recognized with different kind
{
    const syms = extractSymbols('Function Bar() As Long\nEnd Function');
    assert.strictEqual(syms.length, 1, 'one symbol');
    assert.strictEqual(syms[0].name, 'Bar', 'name is Bar');
    assert.strictEqual(syms[0].kind, 12, 'kind 12 = Function');
    assert.ok(syms[0].isFunction === true, 'isFunction flag set');
    console.log('[PASS] Function recognized with kind=12');
}

// 3. Property Get/Let/Set are recognized as kind 6
{
    const code = 'Property Get MyProp() As String\nEnd Property';
    const syms = extractSymbols(code);
    assert.strictEqual(syms.length, 1, 'one symbol');
    assert.strictEqual(syms[0].kind, 6, 'kind 6 = Property');
    assert.ok(syms[0].isProperty === true, 'isProperty flag set');
    console.log('[PASS] Property recognized with kind=6');
}

// 4. Class and its members are collected
{
    const code = `
    Class MyClass
        Public x As Integer
        Public Sub DoSomething()
        End Sub
    End Class
    `;
    const syms = extractSymbols(code);
    assert.strictEqual(syms.length, 1, 'one top-level symbol (the class)');
    const cls = syms[0];
    assert.strictEqual(cls.name, 'MyClass', 'class name');
    assert.strictEqual(cls.kind, 5, 'kind 5 = Class');
    assert.ok(cls.children, 'children array present');
    assert.ok(cls.children.length >= 1, 'class has members');
    console.log('[PASS] Class with members recognized');
}

// 5. Multiple top-level procedures
{
    const code = 'Sub A()\nEnd Sub\nSub B()\nEnd Sub\nFunction C() As Long\nEnd Function';
    const syms = extractSymbols(code);
    assert.strictEqual(syms.length, 3, 'three symbols');
    assert.strictEqual(syms[0].name, 'A', 'first is A');
    assert.strictEqual(syms[1].name, 'B', 'second is B');
    assert.strictEqual(syms[2].name, 'C', 'third is C');
    console.log('[PASS] Multiple procedures extracted');
}

// 6. Global variables are recognized
{
    const code = 'Dim x As Integer\nSub Foo()\nEnd Sub\nDim y As String';
    const syms = extractSymbols(code);
    // Order: x (Dim), Foo (Sub), y (Dim)
    assert.ok(syms.some((s: any) => s.name === 'x' && s.kind === 13), 'global variable x');
    assert.ok(syms.some((s: any) => s.name === 'Foo' && s.kind === 11), 'procedure Foo');
    assert.ok(syms.some((s: any) => s.name === 'y' && s.kind === 13), 'global variable y');
    console.log('[PASS] Global variables recognized');
}

// 7. Empty source produces no symbols
{
    const syms = extractSymbols('');
    assert.strictEqual(syms.length, 0, 'empty source: no symbols');
    console.log('[PASS] Empty source: no symbols');
}

// 8. Range information captures position
{
    const code = 'Sub Bar()\n    x = 1\nEnd Sub';
    const syms = extractSymbols(code);
    assert.strictEqual(syms.length, 1, 'one symbol');
    const range = syms[0].range;
    assert.ok(range.start, 'range.start present');
    assert.strictEqual(range.start.line, 1, 'start line 1');
    assert.strictEqual(range.start.column, 1, 'start column 1');
    assert.ok(range.end, 'range.end present');
    console.log('[PASS] Range information captured');
}

console.log('\n✅ LSP DocumentSymbol: 全テスト通過');
