/**
 * 予約済みステートメントへの引数付与テスト
 *
 * MS-VBAL §5.4.2.12 end-statement   = "End"
 * MS-VBAL §5.4.2.11 stop-statement  = "Stop"
 *
 * どちらも引数を取らない単一キーワード文。
 * MS-VBAL §5.4.1 statement-block = *(block-statement EOS)
 * EOS = *(EOL / ":") であるため、同一行に識別子が続く場合は構文エラーになるべき。
 *
 * 凡例:
 *   [PASS] 正しく構文エラーを報告する
 *   [FAIL] 構文エラーを見逃している（バグ）
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

function assertNoError(label: string, diags: Diag[]) {
    assert.strictEqual(diags.length, 0,
        `${label}: エラーなしを期待 (got: ${diags.map(d => `line=${d.line} "${d.msg}"`).join(', ')})`);
}

// ─── End 文 ──────────────────────────────────────────────────────────────────

// RS01: End Foo — Sub ボディ直下（Ifブロックなし）
// "End" のあと同一行に識別子が続く → 構文エラー（line 2）
{
    const src = [
        'Sub Test()',   // line 1
        '    End Foo', // line 2: 構文エラー
        'End Sub',     // line 3
    ].join('\n');
    const diags = diagnose(src);
    assertError('RS01', diags, 2);
    console.log('[PASS] RS01: End Foo in Sub body → error at line 2');
}

// RS02: End Foo — If ブロック内
// "End" のあと同一行に識別子が続く → 構文エラー（line 4）
{
    const src = [
        'Sub Test()',          // line 1
        '    If x > 0 Then',  // line 2
        '        x = 1',      // line 3
        '    End Foo',        // line 4: 構文エラー
        'End Sub',            // line 5
    ].join('\n');
    const diags = diagnose(src);
    assertError('RS02', diags, 4);
    console.log('[PASS] RS02: End Foo inside If → error at line 4');
}

// RS03: End Foo — For ループ内
{
    const src = [
        'Sub Test()',           // line 1
        '    For i = 1 To 3',  // line 2
        '        End Foo',     // line 3: 構文エラー
        '    Next i',          // line 4
        'End Sub',             // line 5
    ].join('\n');
    const diags = diagnose(src);
    assertError('RS03', diags, 3);
    console.log('[PASS] RS03: End Foo inside For → error at line 3');
}

// RS04: End : Foo — コロン区切りは有効（別文として解析される）
// End は合法な EndStatement、: で区切られた Foo は別文
{
    const src = [
        'Sub Foo()',    // line 1
        'End Sub',     // line 2
        'Sub Test()',   // line 3
        '    End : Foo', // line 4: End は合法、Foo は別の文として続く
        'End Sub',     // line 5
    ].join('\n');
    const diags = diagnose(src);
    assertNoError('RS04', diags);
    console.log('[PASS] RS04: End : Foo (colon separator) → valid, no error');
}

// RS05: End (単独) — 合法な EndStatement
{
    const src = [
        'Sub Test()',  // line 1
        '    End',    // line 2: 合法（プログラム終了文）
        'End Sub',    // line 3
    ].join('\n');
    const diags = diagnose(src);
    assertNoError('RS05', diags);
    console.log('[PASS] RS05: Bare End (alone) → valid EndStatement');
}

// ─── Stop 文 ─────────────────────────────────────────────────────────────────

// RS06: Stop Foo — Sub ボディ直下
// "Stop" のあと同一行に識別子が続く → 構文エラー（line 2）
{
    const src = [
        'Sub Test()',    // line 1
        '    Stop Foo', // line 2: 構文エラー
        'End Sub',      // line 3
    ].join('\n');
    const diags = diagnose(src);
    assertError('RS06', diags, 2);
    console.log('[PASS] RS06: Stop Foo in Sub body → error at line 2');
}

// RS07: Stop Foo — If ブロック内
{
    const src = [
        'Sub Test()',          // line 1
        '    If x > 0 Then',  // line 2
        '        Stop Foo',   // line 3: 構文エラー
        '    End If',         // line 4
        'End Sub',            // line 5
    ].join('\n');
    const diags = diagnose(src);
    assertError('RS07', diags, 3);
    console.log('[PASS] RS07: Stop Foo inside If → error at line 3');
}

// RS08: Stop (単独) — 合法な StopStatement
{
    const src = [
        'Sub Test()',  // line 1
        '    Stop',   // line 2: 合法
        'End Sub',    // line 3
    ].join('\n');
    const diags = diagnose(src);
    assertNoError('RS08', diags);
    console.log('[PASS] RS08: Bare Stop (alone) → valid StopStatement');
}

// RS09: Stop : Foo — コロン区切りは有効
{
    const src = [
        'Sub Foo()',     // line 1
        'End Sub',      // line 2
        'Sub Test()',    // line 3
        '    Stop : Foo', // line 4: Stop は合法、Foo は別の文
        'End Sub',      // line 5
    ].join('\n');
    const diags = diagnose(src);
    assertNoError('RS09', diags);
    console.log('[PASS] RS09: Stop : Foo (colon separator) → valid, no error');
}

console.log('\n✅ 予約済みステートメント引数テスト: 全テスト通過');
