import { Lexer } from '../../src/engine/lexer';
import { Parser } from '../../src/engine/parser';
import { ReferencesProvider } from '../../src/lsp/references-provider';
import { RenameProvider } from '../../src/lsp/rename-provider';
import { assert } from '../../test-libs/test-runner';

const URI = 'file:///test.bas';

function makeProviders(src: string) {
    const tokens = new Lexer(src).tokenize();
    const ast = new Parser(tokens).parse();
    const refs = new ReferencesProvider();
    refs.setDocumentUri(URI);
    const rename = new RenameProvider();
    rename.setDocumentUri(URI);
    return { statements: ast.body, refs, rename };
}

function posOf(src: string, needle: string, occurrence = 1): { line: number; character: number } {
    const lines = src.split('\n');
    let count = 0;
    for (let i = 0; i < lines.length; i++) {
        const idx = lines[i].indexOf(needle);
        if (idx !== -1) {
            count++;
            if (count === occurrence) return { line: i, character: idx };
        }
    }
    throw new Error(`"${needle}" not found (occurrence ${occurrence})`);
}

// 1. Sub の参照箇所をすべて検索（宣言を含む）
{
    const src = [
        'Sub CalcTotal()',
        '    Dim x As Long',
        'End Sub',
        'Sub Main()',
        '    CalcTotal',
        '    CalcTotal',
        'End Sub',
    ].join('\n');

    const { statements, refs } = makeProviders(src);
    const pos = posOf(src, 'CalcTotal');
    const result = refs.getReferences(statements, src, pos.line, pos.character, true);
    assert.strictEqual(result.length, 3, 'CalcTotal: 宣言+2呼び出し = 3件');
    console.log('[PASS] Sub 参照 (includeDeclaration=true): 3件');
}

// 2. 宣言を除いた参照（includeDeclaration=false）
{
    const src = [
        'Sub CalcTotal()',
        'End Sub',
        'Sub Main()',
        '    CalcTotal',
        'End Sub',
    ].join('\n');

    const { statements, refs } = makeProviders(src);
    const pos = posOf(src, 'CalcTotal');
    const result = refs.getReferences(statements, src, pos.line, pos.character, false);
    assert.strictEqual(result.length, 1, '宣言除外で呼び出し箇所のみ 1件');
    console.log('[PASS] includeDeclaration=false: 呼び出し箇所のみ');
}

// 3. 変数の参照
{
    const src = [
        'Function Foo()',
        '    Dim total As Long',
        '    total = 10',
        '    Foo = total + 1',
        'End Function',
    ].join('\n');

    const { statements, refs } = makeProviders(src);
    const pos = posOf(src, 'total');
    const result = refs.getReferences(statements, src, pos.line, pos.character, true);
    assert.strictEqual(result.length, 3, 'total: 宣言+代入+参照 = 3件');
    console.log('[PASS] 変数参照 3件');
}

// 4. コメント行・文字列リテラルは除外される
{
    const src = [
        "' CalcTotal is a function",
        'Sub CalcTotal()',
        '    Dim s As String',
        '    s = "CalcTotal result"',
        'End Sub',
        'Sub Main()',
        '    CalcTotal',
        'End Sub',
    ].join('\n');

    const { statements, refs } = makeProviders(src);
    const pos = posOf(src, 'CalcTotal', 2); // 2番目の CalcTotal（Sub宣言）
    const result = refs.getReferences(statements, src, pos.line, pos.character, true);
    // コメント行と文字列リテラルの CalcTotal は除外
    assert.strictEqual(result.length, 2, 'コメント・文字列除外で 宣言+呼び出し = 2件');
    console.log('[PASS] コメント・文字列リテラルを除外: 2件');
}

// 5. 行末コメント中の参照は除外
{
    const src = [
        'Sub CalcTotal()',
        'End Sub',
        'Sub Main()',
        "    CalcTotal ' call CalcTotal here",
        'End Sub',
    ].join('\n');

    const { statements, refs } = makeProviders(src);
    const pos = posOf(src, 'CalcTotal');
    const result = refs.getReferences(statements, src, pos.line, pos.character, true);
    assert.strictEqual(result.length, 2, '行末コメント内は除外: 宣言+呼び出し = 2件');
    console.log('[PASS] 行末コメント除外: 2件');
}

// 6. カーソルが ')' 上（識別子なし） → 空配列
{
    const src = 'Sub Foo()\nEnd Sub\n';
    const { statements, refs } = makeProviders(src);
    // "Sub Foo()" の ')' は index 8 — 識別子文字でない
    const result = refs.getReferences(statements, src, 0, 8, true);
    assert.strictEqual(result.length, 0, ') 上のカーソル → 空配列');
    console.log('[PASS] 識別子でない位置: 空配列');
}

// 7. Rename: 全参照を新しい名前に置換
{
    const src = [
        'Sub CalcTotal()',
        'End Sub',
        'Sub Main()',
        '    CalcTotal',
        '    CalcTotal',
        'End Sub',
    ].join('\n');

    const { statements, rename } = makeProviders(src);
    const pos = posOf(src, 'CalcTotal');
    const edits = rename.getRename(statements, src, pos.line, pos.character, 'ComputeTotal');
    assert.ok(edits !== null, 'edits が null でない');
    assert.strictEqual(edits!.length, 3, 'CalcTotal: 3件のテキスト編集');
    assert.strictEqual(edits![0].newText, 'ComputeTotal', '新しい名前が正しい');
    console.log('[PASS] Rename: 3件のテキスト編集');
}

// 8. Rename: 範囲が正確
{
    const src = [
        'Sub CalcTotal()',
        'End Sub',
    ].join('\n');

    const { statements, rename } = makeProviders(src);
    const pos = posOf(src, 'CalcTotal');
    const edits = rename.getRename(statements, src, pos.line, pos.character, 'X');
    assert.ok(edits !== null && edits.length > 0, 'edits あり');
    const e = edits![0];
    assert.strictEqual(e.range.start.line, 0, '行 0');
    assert.strictEqual(e.range.start.character, 4, '列 4（"Sub " の後）');
    assert.strictEqual(e.range.end.character, 4 + 'CalcTotal'.length, '終端列');
    console.log('[PASS] Rename: 範囲が正確');
}

// 9. 大文字小文字を区別しない検索
{
    const src = [
        'Sub calcTotal()',
        'End Sub',
        'Sub Main()',
        '    CALCTOTAL',
        '    CalcTotal',
        'End Sub',
    ].join('\n');

    const { statements, refs } = makeProviders(src);
    const pos = posOf(src, 'calcTotal');
    const result = refs.getReferences(statements, src, pos.line, pos.character, true);
    assert.strictEqual(result.length, 3, '大文字小文字を区別しない: 3件');
    console.log('[PASS] 大文字小文字を区別しない: 3件');
}

console.log('\n✅ LSP References / Rename: 全テスト通過');
