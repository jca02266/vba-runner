import { Lexer, TokenType, VBA_CONTEXTUAL_KEYWORDS, VBA_KEYWORD_CATEGORIES } from '../../src/engine/lexer';
import { assert } from '../../test-libs/test-runner';

function token(word: string): TokenType {
    return new Lexer(word).tokenize()[0].type;
}

// The spelling table is metadata; tokenization remains unchanged.
assert.deepStrictEqual(VBA_KEYWORD_CATEGORIES.new,
    ['marker-keyword', 'operator-identifier']);
assert.deepStrictEqual(VBA_KEYWORD_CATEGORIES.case,
    ['statement-keyword', 'marker-keyword']);
assert.deepStrictEqual(VBA_KEYWORD_CATEGORIES.long,
    ['reserved-type-identifier']);
assert.deepStrictEqual(VBA_KEYWORD_CATEGORIES.true,
    ['literal-identifier']);

// Contextual spellings keep their dedicated lexer tokens for parser context.
for (const [word, expected] of [
    ['Step', TokenType.KeywordStep],
    ['Output', TokenType.KeywordOutput],
    ['Class', TokenType.KeywordClass],
    ['Property', TokenType.KeywordProperty],
    ['Error', TokenType.KeywordError],
] as const) {
    assert.strictEqual(token(word), expected, `${word} token type is stable`);
    assert.deepStrictEqual(VBA_KEYWORD_CATEGORIES[word.toLowerCase()], [],
        `${word} is contextual, not a reserved category`);
    assert.strictEqual(VBA_CONTEXTUAL_KEYWORDS.has(word.toLowerCase()), true,
        `${word} is listed as contextual`);
}

// Reserved keywords remain reserved at the lexer level.
assert.strictEqual(token('Shared'), TokenType.KeywordShared);
assert.strictEqual(token('Function'), TokenType.KeywordFunction);

console.log('[PASS] VBA keyword category metadata and token stability');
