import { Lexer } from '../../src/engine/lexer';
import { Parser } from '../../src/engine/parser';
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

// ─── scope / Const 表示のレグレッション ────────────────────────────────────────
// [回帰] Bug E2: scope が小文字で表示されていた
// [回帰] Bug E1: Const に型・値が表示されなかった

{
    // Public 変数は "Public" と大文字で表示される
    const code = ['Public count As Long', 'Sub Test()', '  count = 1', 'End Sub'].join('\n');
    const info = getHoverAt(code, 2, 3); // 'count' on usage line
    assert.ok(info?.contents.includes('Public count As Long'), `Public scope capitalized: ${info?.contents}`);
    console.log('[PASS] Public変数 hover: scope 大文字');
}

{
    // Private Function は "Private Function" と表示される
    const code = ['Private Function Calc() As Long', '  Calc = 1', 'End Function'].join('\n');
    const info = getHoverAt(code, 0, 18); // 'Calc'
    assert.ok(info?.contents.includes('Private Function Calc'), `Private scope capitalized: ${info?.contents}`);
    console.log('[PASS] Private Function hover: scope 大文字');
}

{
    // Const は型名と値を含んで表示される
    const code = ['Const MAX As Long = 100', 'Sub Test()', '  Debug.Print MAX', 'End Sub'].join('\n');
    const info = getHoverAt(code, 2, 14); // 'MAX' on usage line
    assert.ok(info?.contents.includes('Const MAX As Long = 100'), `Const type+value: ${info?.contents}`);
    console.log('[PASS] Const hover: 型と値を表示');
}

{
    // 型なし Const は値のみ表示される
    const code = ['Const PI = 3.14', 'Sub Test()', '  Debug.Print PI', 'End Sub'].join('\n');
    const info = getHoverAt(code, 2, 14); // 'PI' on usage line
    assert.ok(info?.contents.includes('Const PI = 3.14'), `Const no-type value: ${info?.contents}`);
    console.log('[PASS] Const hover: 型なしでも値を表示');
}

// ─── UDT (Type) / Enum member ホバーのレグレッション ──────────────────────────
// [回帰] Bug LSP-2: Enum 定数のホバーが null だった
// [回帰] Bug LSP-3: UDT (Type) 名のホバーが null だった

{
    const code = [
        'Type Point',
        '    X As Long',
        '    Y As Long',
        'End Type',
        'Enum Direction',
        '    North = 1',
        '    South = 2',
        'End Enum',
        'Sub Test()',
        '    Dim pt As Point',
        '    Dim d As Direction',
        '    d = North',
        'End Sub',
    ].join('\n');

    // UDT 型名ホバー (line 0 col 5 → "Point")
    const hoverPoint = getHoverAt(code, 0, 5);
    assert.ok(hoverPoint !== null, 'UDT Type名 hover は null でない');
    assert.ok(hoverPoint?.contents.includes('Type Point'), `UDT hover に Type Point を含む: ${hoverPoint?.contents}`);
    console.log('[PASS] UDT (Type) 名 hover: Type Point を表示（Bug LSP-3）');

    // Enum 定数ホバー (line 11 col 8 → "North")
    const hoverNorth = getHoverAt(code, 11, 8);
    assert.ok(hoverNorth !== null, 'Enum定数 hover は null でない');
    assert.ok(hoverNorth?.contents.includes('Direction.North'), `Enum hover に Direction.North を含む: ${hoverNorth?.contents}`);
    console.log('[PASS] Enum 定数 hover: Direction.North = 1 を表示（Bug LSP-2）');
}

console.log('\n✅ LSP Hover: 全テスト通過');
