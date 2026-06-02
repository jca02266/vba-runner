/**
 * ブロック終了キーワードのタイポ検出テスト
 *
 * If / For ブロックの閉じキーワードを誤って書いた場合のエラー位置を検証する。
 * MS-VBAL §5.4.2.1: statement-block = *(block-statement EOS)
 * MS-VBAL §5.4.2.8: if-statement requires "End" "If" to close
 * MS-VBAL §5.4.2.3: for-statement requires "Next" to close
 *
 * 凡例:
 *   [PASS] タイポのある行でエラーが報告される（正しい動作）
 *   [FAIL] エラー位置がずれている（バグ）
 */

import { Lexer, LexError } from '../../src/engine/lexer';
import { Parser } from '../../src/engine/parser';
import { assert } from '../../test-libs/test-runner';

interface Diag { line: number; col: number; msg: string; src: 'lexer' | 'parser' }

function diagnose(src: string): Diag[] {
    const lexer = new Lexer(src);
    let tokens;
    try {
        tokens = lexer.tokenize();
    } catch (e) {
        if (e instanceof LexError) {
            return [{ line: e.line, col: e.column, msg: e.message.replace(/ \(line \d+\)$/, ''), src: 'lexer' }];
        }
        throw e;
    }
    const ast = new Parser(tokens, { errorRecovery: true }).parse();
    return [
        ...lexer.diagnostics.map(d => ({ line: d.line, col: d.column, msg: d.message, src: 'lexer' as const })),
        ...ast.diagnostics.map(d => ({ line: d.loc.start.line, col: d.loc.start.column, msg: d.message, src: 'parser' as const })),
    ];
}

function assertError(label: string, diags: Diag[], expectedLine: number, msgFragment?: string) {
    assert.ok(diags.length >= 1, `${label}: エラーが1件以上あること`);
    const d = diags[0];
    assert.strictEqual(d.line, expectedLine,
        `${label}: エラーが line=${expectedLine} を指すこと (got line=${d.line}, msg="${d.msg}")`);
    if (msgFragment) {
        assert.ok(d.msg.includes(msgFragment),
            `${label}: メッセージに "${msgFragment}" を含むこと (got: "${d.msg}")`);
    }
}

// ─── If ブロックの閉じキーワードタイポ ──────────────────────────────────────

// BT01: End Sub as typo for End If
// If の中を End Sub で閉じようとした場合、End Sub の行でエラーになる
{
    const src = [
        'Sub Test()',          // line 1
        '    If x > 0 Then',  // line 2
        '        x = 2',      // line 3
        '    End Sub',        // line 4: タイポ（End If が正しい）
        'End Sub',            // line 5
    ].join('\n');
    const diags = diagnose(src);
    assertError('BT01', diags, 4);
    console.log('[PASS] BT01: End Sub as End If typo → error at line 4');
}

// BT02: End Foo as typo for End If
// End <識別子> は EOS なしで End の後に続くため構文エラー
// エラーは End Foo のある行（line 4）で報告されるべき
{
    const src = [
        'Sub Test()',          // line 1
        '    If x > 0 Then',  // line 2
        '        x = 2',      // line 3
        '    End Foo',        // line 4: タイポ（End If が正しい）
        'End Sub',            // line 5
    ].join('\n');
    const diags = diagnose(src);
    assertError('BT02', diags, 4);
    console.log('[PASS] BT02: End Foo as End If typo → error at line 4');
}

// BT03: End Sub as typo for Next (For ループ)
{
    const src = [
        'Sub Test()',           // line 1
        '    For i = 1 To 3',  // line 2
        '        x = i',       // line 3
        '    End Sub',         // line 4: タイポ（Next が正しい）
        'End Sub',             // line 5
    ].join('\n');
    const diags = diagnose(src);
    assertError('BT03', diags, 4);
    console.log('[PASS] BT03: End Sub as Next typo in For → error at line 4');
}

// BT04: End Foo as typo for Next (For ループ)
// エラーは End Foo のある行（line 4）で報告されるべき
{
    const src = [
        'Sub Test()',           // line 1
        '    For i = 1 To 3',  // line 2
        '        x = i',       // line 3
        '    End Foo',         // line 4: タイポ（Next が正しい）
        'End Sub',             // line 5
    ].join('\n');
    const diags = diagnose(src);
    assertError('BT04', diags, 4);
    console.log('[PASS] BT04: End Foo as Next typo in For → error at line 4');
}

// BT05: End Sub as typo for Loop (Do...Loop)
{
    const src = [
        'Sub Test()',            // line 1
        '    Do While x > 0',   // line 2
        '        x = x - 1',   // line 3
        '    End Sub',          // line 4: タイポ（Loop が正しい）
        'End Sub',              // line 5
    ].join('\n');
    const diags = diagnose(src);
    assertError('BT05', diags, 4);
    console.log('[PASS] BT05: End Sub as Loop typo in Do → error at line 4');
}

// BT06: End Foo as typo for Loop (Do...Loop)
{
    const src = [
        'Sub Test()',            // line 1
        '    Do While x > 0',   // line 2
        '        x = x - 1',   // line 3
        '    End Foo',          // line 4: タイポ（Loop が正しい）
        'End Sub',              // line 5
    ].join('\n');
    const diags = diagnose(src);
    assertError('BT06', diags, 4);
    console.log('[PASS] BT06: End Foo as Loop typo in Do → error at line 4');
}

// BT07: End Sub as typo for Wend (While...Wend)
{
    const src = [
        'Sub Test()',            // line 1
        '    While x > 0',      // line 2
        '        x = x - 1',   // line 3
        '    End Sub',          // line 4: タイポ（Wend が正しい）
        'End Sub',              // line 5
    ].join('\n');
    const diags = diagnose(src);
    assertError('BT07', diags, 4);
    console.log('[PASS] BT07: End Sub as Wend typo in While → error at line 4');
}

// BT08: End Foo as typo for Wend (While...Wend)
{
    const src = [
        'Sub Test()',            // line 1
        '    While x > 0',      // line 2
        '        x = x - 1',   // line 3
        '    End Foo',          // line 4: タイポ（Wend が正しい）
        'End Sub',              // line 5
    ].join('\n');
    const diags = diagnose(src);
    assertError('BT08', diags, 4);
    console.log('[PASS] BT08: End Foo as Wend typo in While → error at line 4');
}

console.log('\n✅ ブロック終了タイポテスト: 全テスト通過');
