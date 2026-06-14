import { Lexer } from '../../src/engine/lexer';
import { Parser } from '../../src/engine/parser';
import { CodeLensProvider } from '../../src/lsp/code-lens-provider';
import { assert } from '../../test-libs/test-runner';

const URI = 'file:///test.bas';

function getCodeLens(src: string) {
    const tokens = new Lexer(src).tokenize();
    const ast = new Parser(tokens).parse();
    const provider = new CodeLensProvider();
    return provider.getCodeLens(ast.body, src, URI);
}

function lensesFor(items: any[], procLine: number) {
    return items.filter(i => i.range.start.line === procLine);
}

// 1. 引数なし Sub に ▶ Run が付く
{
    const src = ['Sub Foo()', 'End Sub'].join('\n');
    const items = getCodeLens(src);
    const runLens = items.find((i: any) => i.command.title === '▶ Run');
    assert.ok(runLens, '▶ Run が存在する');
    assert.strictEqual(runLens.command.command, 'vba-runner.runProcedure', 'コマンド名');
    assert.strictEqual(runLens.command.arguments[1], 'Foo', 'プロシージャ名');
    console.log('[PASS] 引数なし Sub に ▶ Run');
}

// 2. 必須引数ありの Sub に ▶ Run は付かない
{
    const src = ['Sub Bar(x As Long)', 'End Sub'].join('\n');
    const items = getCodeLens(src);
    const runLens = items.find((i: any) => i.command.title === '▶ Run');
    assert.ok(!runLens, '必須引数ありは ▶ Run なし');
    console.log('[PASS] 必須引数あり Sub に ▶ Run なし');
}

// 3. Optional のみの Sub には ▶ Run が付く
{
    const src = ['Sub Baz(Optional x As Long = 0)', 'End Sub'].join('\n');
    const items = getCodeLens(src);
    const runLens = items.find((i: any) => i.command.title === '▶ Run');
    assert.ok(runLens, 'Optional のみなら ▶ Run あり');
    console.log('[PASS] Optional のみ Sub に ▶ Run');
}

// 4. 参照数が正しく表示される
{
    const src = [
        'Sub CalcTotal()',
        'End Sub',
        'Sub Main()',
        '    CalcTotal',
        '    CalcTotal',
        'End Sub',
    ].join('\n');
    const items = getCodeLens(src);
    const refLens = items.find((i: any) => i.command.title.includes('references') || i.command.title.includes('reference'));
    assert.ok(refLens, '参照数レンズが存在する');
    assert.ok(refLens.command.title.includes('2'), '参照数 2 が表示される');
    console.log('[PASS] 参照数 2:', refLens.command.title);
}

// 5. 参照0の Private Sub に警告マーク
{
    const src = [
        'Private Sub UnusedHelper()',
        'End Sub',
        'Sub Main()',
        'End Sub',
    ].join('\n');
    const items = getCodeLens(src);
    const line0lenses = lensesFor(items, 0);
    const refLens = line0lenses.find((i: any) => i.command.title.includes('0 references'));
    assert.ok(refLens, '0 references レンズが存在する');
    assert.ok(refLens.command.title.startsWith('⚠'), 'Private + 0 refs に ⚠');
    console.log('[PASS] Private 0-references に ⚠:', refLens.command.title);
}

// 6. 参照0の Public Sub には警告マークなし（エントリーポイント候補）
{
    const src = [
        'Sub EntryPoint()',
        'End Sub',
    ].join('\n');
    const items = getCodeLens(src);
    const refLens = items.find((i: any) => i.command.title.includes('0 references'));
    assert.ok(refLens, '0 references レンズが存在する');
    assert.ok(!refLens.command.title.startsWith('⚠'), 'Public の 0 refs に ⚠ なし');
    console.log('[PASS] Public 0-references に ⚠ なし:', refLens.command.title);
}

// 7. Test_ プロシージャで参照されると「テスト済み」になる
{
    const src = [
        'Function CalcTotal()',
        '    CalcTotal = 42',
        'End Function',
        'Sub Test_CalcTotal()',
        '    Dim r',
        '    r = CalcTotal()',
        'End Sub',
    ].join('\n');
    const items = getCodeLens(src);
    const line0lenses = lensesFor(items, 0);
    const testLens = line0lenses.find((i: any) =>
        i.command.title === '✓ Tested' || i.command.title === 'Untested'
    );
    assert.ok(testLens, 'テスト状態レンズが存在する');
    assert.strictEqual(testLens.command.title, '✓ Tested', 'Test_ から参照されると Tested');
    console.log('[PASS] Test_ から参照 → Tested');
}

// 8. Test_ から参照されなければ「未テスト」
{
    const src = [
        'Function CalcTotal()',
        '    CalcTotal = 42',
        'End Function',
        'Sub Main()',
        '    CalcTotal',
        'End Sub',
    ].join('\n');
    const items = getCodeLens(src);
    const line0lenses = lensesFor(items, 0);
    const testLens = line0lenses.find((i: any) =>
        i.command.title === '✓ Tested' || i.command.title === 'Untested'
    );
    assert.strictEqual(testLens?.command.title, 'Untested', 'Test_ なし → Untested');
    console.log('[PASS] Test_ なし → Untested');
}

// 9. 複数プロシージャにそれぞれレンズが付く
{
    const src = [
        'Sub Foo()',
        'End Sub',
        'Sub Bar()',
        'End Sub',
        'Sub Baz()',
        'End Sub',
    ].join('\n');
    const items = getCodeLens(src);
    const lines = new Set(items.map((i: any) => i.range.start.line));
    assert.ok(lines.has(0), 'Foo の行 0 にレンズあり');
    assert.ok(lines.has(2), 'Bar の行 2 にレンズあり');
    assert.ok(lines.has(4), 'Baz の行 4 にレンズあり');
    console.log('[PASS] 複数プロシージャにそれぞれレンズ');
}

console.log('\n✅ Code Lens: 全テスト通過');
