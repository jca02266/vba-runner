import { assert } from '../../test-libs/test-runner';
import { needsLineContinuation } from '../../src/lsp/line-continuation-checker';

// ─── 継続が必要なケース ──────────────────────────────────────────────────────

const shouldContinue: [string, string][] = [
    ['    x = 1 +',       '+ で終わる'],
    ['    x = 1 -',       '- で終わる'],
    ['    x = a *',       '* で終わる'],
    ['    x = a /',       '/ で終わる'],
    ['    x = a \\',      '\\ で終わる'],
    ['    x = a ^',       '^ で終わる'],
    ['    x = a &',       '& で終わる'],
    ['    Foo(a,',        ', で終わる'],
    ['    x = Foo(',      '( で終わる'],
    ['    Set ws =',      '= で終わる'],
    ['    If x <',        '< で終わる'],
    ['    If x >',        '> で終わる'],
    ['    If x <=',       '<= で終わる'],
    ['    If x >=',       '>= で終わる'],
    ['    If x <>',       '<> で終わる'],
    ['    If x And',      'And で終わる'],
    ['    If x Or',       'Or で終わる'],
    ['    result = a Xor','Xor で終わる'],
    ['    result = a Eqv','Eqv で終わる'],
    ['    result = a Imp','Imp で終わる'],
    ['    n = x Mod',     'Mod で終わる'],
    ['    If x Like',     'Like で終わる'],
    ['    If x Is',       'Is で終わる'],
];

for (const [line, desc] of shouldContinue) {
    assert.strictEqual(needsLineContinuation(line), true, `継続必要: ${desc}`);
    console.log(`[PASS] 継続必要: ${desc}`);
}

// ─── 継続が不要なケース ──────────────────────────────────────────────────────

const shouldNotContinue: [string, string][] = [
    ['    x = 1 + _',          'すでに _ がある'],
    ["    ' コメント行",        'コメント行'],
    ['',                       '空行'],
    ['    MsgBox "hello"',     '完結した文'],
    ['    End Sub',            'End Sub'],
    ['    Next i',             'Next i'],
    ['    Loop',               'Loop'],
    ['    End If',             'End If'],
    ['    Debug.Print x',      'メソッド呼び出し'],
    ['    x = 1',              '完結した代入'],
];

for (const [line, desc] of shouldNotContinue) {
    assert.strictEqual(needsLineContinuation(line), false, `継続不要: ${desc}`);
    console.log(`[PASS] 継続不要: ${desc}`);
}

// ─── 大文字小文字の扱い ──────────────────────────────────────────────────────

{
    assert.strictEqual(needsLineContinuation('    If x and'), true, 'and (小文字) → 継続必要');
    console.log('[PASS] and 小文字: 継続必要');
}
{
    assert.strictEqual(needsLineContinuation('    If x AND'), true, 'AND (大文字) → 継続必要');
    console.log('[PASS] AND 大文字: 継続必要');
}

// ─── インラインコメントがある行は継続不要（コメント後の _ は無効） ─────────────

const inlineCommentCases: [string, string][] = [
    ['    Array(1, 2), \' コメント',          ', の後にコメント'],
    ['    x = 1 + \' note',                   '+ の後にコメント'],
    ['    If x And \' check',                 'And の後にコメント'],
    ['    x = "hello" \' use Or Or Or',       'コメント内にキーワード Or'],
    ['    Foo(a, b) \' done',                 '完結した呼び出し+コメント'],
];

for (const [line, desc] of inlineCommentCases) {
    assert.strictEqual(needsLineContinuation(line), false, `継続不要: ${desc}`);
    console.log(`[PASS] 継続不要（インラインコメント）: ${desc}`);
}

console.log('\n✅ needsLineContinuation: 全テスト通過');
