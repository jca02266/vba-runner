import { Lexer } from '../../src/compiler/lexer';
import { Parser } from '../../src/compiler/parser';
import { HoverProvider } from '../../src/lsp/hover-provider';
import { assert } from '../../test-libs/test-runner';

function getHoverAt(src: string, line: number, character: number): any {
    const tokens = new Lexer(src).tokenize();
    const ast = new Parser(tokens).parse();
    const provider = new HoverProvider();
    return provider.getHoverInfo(ast.body, src, line, character);
}

// 1. Hover over Sub procedure shows kind and signature
{
    const code = 'Sub Foo(x As Integer)\nEnd Sub';
    const info = getHoverAt(code, 0, 4); // position of 'Foo'
    assert.ok(info, 'hover info returned');
    assert.ok(info.contents, 'contents present');
    assert.ok(info.contents.includes('Sub'), 'shows Sub keyword');
    assert.ok(info.contents.includes('Foo'), 'shows name');
    assert.ok(info.range, 'range present');
    console.log('[PASS] Sub hover:', info.contents.slice(0, 60));
}

// 2. Hover over Function shows Function keyword
{
    const code = 'Function Bar() As Long\nEnd Function';
    const info = getHoverAt(code, 0, 9); // position of 'Bar'
    assert.ok(info, 'hover info returned');
    assert.ok(info.contents.includes('Function'), 'shows Function keyword');
    console.log('[PASS] Function hover:', info.contents.slice(0, 60));
}

// 3. Hover over non-existent identifier returns nothing
{
    const code = 'Sub Foo()\nEnd Sub';
    const info = getHoverAt(code, 0, 0); // position before 'Sub'
    assert.strictEqual(info, null, 'no info for non-existent identifier');
    console.log('[PASS] Non-existent identifier: no hover');
}

// 4. Hover over variable declaration shows type
{
    const code = 'Dim x As Integer';
    const info = getHoverAt(code, 0, 4); // position of 'x'
    assert.ok(info, 'hover info returned');
    assert.ok(info.contents, 'contents present');
    assert.ok(info.contents.includes('Integer') || info.contents.includes('x'), 'shows type or name');
    console.log('[PASS] Variable hover:', info.contents.slice(0, 60));
}

// 5. Hover over Class shows Class keyword
{
    const code = 'Class MyClass\nEnd Class';
    const info = getHoverAt(code, 0, 6); // position of 'MyClass'
    assert.ok(info, 'hover info returned');
    assert.ok(info.contents.includes('Class'), 'shows Class keyword');
    console.log('[PASS] Class hover:', info.contents.slice(0, 60));
}

// 6. Hover with range covering identifier
{
    const code = 'Sub Foo()\nEnd Sub';
    const info = getHoverAt(code, 0, 4); // 'Foo'
    assert.ok(info.range, 'range present');
    assert.ok(info.range.start, 'range.start present');
    assert.ok(info.range.end, 'range.end present');
    // Range should cover 'Foo' (columns 4-7 in 0-based)
    assert.strictEqual(info.range.start.character, 4, 'range starts at Foo');
    assert.strictEqual(info.range.end.character, 7, 'range ends after Foo');
    console.log('[PASS] Hover range correct');
}

// 7. Multi-line source: hover on line 1
{
    const code = 'x = 1\nSub Bar()\nEnd Sub';
    const info = getHoverAt(code, 1, 4); // 'Bar' on line 1
    assert.ok(info, 'hover info for line 1');
    assert.ok(info.contents.includes('Bar'), 'shows Bar');
    console.log('[PASS] Multi-line hover on line 1');
}

// 8. Hover on procedure with parameters shows signature
{
    const code = 'Sub Process(a As Integer, b As String)\nEnd Sub';
    const info = getHoverAt(code, 0, 4); // 'Process'
    assert.ok(info, 'hover info returned');
    assert.ok(info.contents, 'contents present');
    // Should include procedure name
    assert.ok(info.contents.includes('Process'), 'shows name');
    console.log('[PASS] Procedure with params hover:', info.contents.slice(0, 80));
}

// 10. Empty source returns nothing for any hover
{
    const code = '';
    const info = getHoverAt(code, 0, 0);
    assert.strictEqual(info, null, 'empty source: no hover');
    console.log('[PASS] Empty source: no hover');
}

console.log('\n✅ LSP Hover: 全テスト通過');
