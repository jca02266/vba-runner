import { assert } from '../../test-libs/test-runner';
import { canonicalKeyword, isInStringOrComment } from '../../src/lsp/keyword-casing';

// ─── 小文字・大文字混在 → 正規ケースに補正 ──────────────────────────────────

const shouldCorrect: [string, string][] = [
    ['if', 'If'],
    ['IF', 'If'],
    ['then', 'Then'],
    ['elseif', 'ElseIf'],
    ['withevents', 'WithEvents'],
    ['paramarray', 'ParamArray'],
    ['raiseevent', 'RaiseEvent'],
    ['addressof', 'AddressOf'],
    ['typeof', 'TypeOf'],
    ['goto', 'GoTo'],
    ['gosub', 'GoSub'],
    ['ptrsafe', 'PtrSafe'],
    ['byref', 'ByRef'],
    ['byval', 'ByVal'],
    ['dim', 'Dim'],
    ['redim', 'ReDim'],
    ['long', 'Long'],
    ['STRING', 'String'],
    ['boolean', 'Boolean'],
    ['nothing', 'Nothing'],
    ['true', 'True'],
];

for (const [input, expected] of shouldCorrect) {
    assert.strictEqual(canonicalKeyword(input), expected, `${input} → ${expected}`);
    console.log(`[PASS] 補正: ${input} → ${expected}`);
}

// ─── すでに正規ケース → undefined（no-op 編集を避ける） ──────────────────────

const alreadyCanonical = ['If', 'WithEvents', 'ElseIf', 'Dim', 'Long', 'ByRef'];
for (const word of alreadyCanonical) {
    assert.strictEqual(canonicalKeyword(word), undefined, `${word} はすでに正規 → undefined`);
    console.log(`[PASS] 正規ケース据え置き: ${word}`);
}

// ─── キーワードでない語 → undefined ─────────────────────────────────────────

const notKeywords = ['foo', 'iffy', 'myIf', 'counter', 'x', 'doStuff'];
for (const word of notKeywords) {
    assert.strictEqual(canonicalKeyword(word), undefined, `${word} は非キーワード → undefined`);
    console.log(`[PASS] 非キーワード: ${word}`);
}

// ─── 文字列・コメント判定 ────────────────────────────────────────────────────

// "...col..." の col が文字列/コメント内かを判定
const stringComment: [string, number, boolean, string][] = [
    ['x = "if"', 6, true, '文字列リテラル内の if'],
    ['x = "if"', 0, false, '文字列の外（先頭）'],
    ["' if then", 3, true, 'コメント内の if'],
    ['Dim x ' + "' if", 8, true, '行末コメント内'],
    ['If x Then', 0, false, '通常コード（先頭）'],
    ['If x Then', 5, false, '通常コード（Then 位置）'],
    ['s = "a" & if', 11, false, '文字列を閉じた後の if'],
];

for (const [line, col, expected, desc] of stringComment) {
    assert.strictEqual(isInStringOrComment(line, col), expected, desc);
    console.log(`[PASS] 文字列/コメント判定: ${desc}`);
}

console.log('\n✅ keyword-casing: 全テスト通過');
