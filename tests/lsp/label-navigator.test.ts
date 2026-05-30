import * as assert from 'assert';
import { Lexer } from '../../src/engine/lexer';
import { Parser } from '../../src/engine/parser';
import { findLabelDefinition, findGoToReferences, isOnLabel } from '../../src/lsp/label-navigator';

const URI = 'file:///test.bas';

function parse(src: string) {
    const tokens = new Lexer(src).tokenize();
    return new Parser(tokens, { errorRecovery: true }).parse().body;
}

// ─── findLabelDefinition ────────────────────────────────────────────────────

{
    const src = [
        'Sub Test()',
        '    If x = "" Then GoTo Skip',
        '    Debug.Print x',
        'Skip:',
        '    Debug.Print "done"',
        'End Sub',
    ].join('\n');
    const stmts = parse(src);

    // GoTo のラベル名 "Skip" 上（line=1, GoTo のあと "Skip" は col 24〜27 くらい）
    // line は 0-based, GoTo は "    If x = "" Then GoTo Skip"
    // "GoTo " = 5文字、"GoTo " starts at col 19 (0-based), label at col 24
    const def = findLabelDefinition(stmts, 1, 24, URI);
    assert.ok(def !== null, 'GoTo上でラベル定義が見つかること');
    assert.strictEqual(def!.range.start.line, 3, 'ラベル定義の行（0-based line 3: "Skip:"）');
    console.log('[PASS] findLabelDefinition: GoTo → LabelStatement');
}

{
    const src = [
        'Sub Test()',
        '    GoTo MyLabel',
        'MyLabel:',
        'End Sub',
    ].join('\n');
    const stmts = parse(src);

    // GoTo 上でない位置（line=2, ラベル定義行）
    const def = findLabelDefinition(stmts, 2, 0, URI);
    assert.strictEqual(def, null, 'LabelStatement上では定義ジャンプは null');
    console.log('[PASS] findLabelDefinition: LabelStatement上では null');
}

// ─── findGoToReferences ─────────────────────────────────────────────────────

{
    const src = [
        'Sub Test()',
        '    For i = 1 To 10',
        '        If i = 5 Then GoTo Continue',
        '        Debug.Print i',
        'Continue:',
        '    Next i',
        'End Sub',
    ].join('\n');
    const stmts = parse(src);

    // LabelStatement "Continue" 上（line=4, col=0）
    const refs = findGoToReferences(stmts, 4, 0, URI, false);
    assert.strictEqual(refs.length, 1, 'GoTo参照が1件見つかること');
    assert.strictEqual(refs[0].range.start.line, 2, 'GoTo参照の行（0-based line 2）');
    console.log('[PASS] findGoToReferences: LabelStatement → GoTo 一覧');
}

{
    // 複数の GoTo が同じラベルを参照
    const src = [
        'Sub Test()',
        '    GoTo Done',
        '    GoTo Done',
        'Done:',
        'End Sub',
    ].join('\n');
    const stmts = parse(src);

    const refs = findGoToReferences(stmts, 3, 0, URI, false);
    assert.strictEqual(refs.length, 2, '複数のGoTo参照が見つかること');
    console.log('[PASS] findGoToReferences: 複数GoTo参照');
}

{
    // includeDeclaration = true のとき LabelStatement 自身も含む
    const src = [
        'Sub Test()',
        '    GoTo Target',
        'Target:',
        'End Sub',
    ].join('\n');
    const stmts = parse(src);

    const refs = findGoToReferences(stmts, 2, 0, URI, true);
    assert.strictEqual(refs.length, 2, 'includeDeclaration=true で定義含め2件');
    console.log('[PASS] findGoToReferences: includeDeclaration=true');
}

// ─── isOnLabel ──────────────────────────────────────────────────────────────

{
    const src = [
        'Sub Test()',
        '    GoTo MyLabel',
        'MyLabel:',
        'End Sub',
    ].join('\n');
    const stmts = parse(src);

    // GoTo のラベル名上
    assert.strictEqual(isOnLabel(stmts, 1, 10), true, 'GoTo ラベル名上は true');
    // LabelStatement 上
    assert.strictEqual(isOnLabel(stmts, 2, 0), true, 'LabelStatement 上は true');
    // 無関係な位置
    assert.strictEqual(isOnLabel(stmts, 0, 0), false, '無関係な位置は false');
    console.log('[PASS] isOnLabel');
}
