import { Lexer } from '../../src/compiler/lexer';
import { Parser } from '../../src/compiler/parser';
import { Evaluator } from '../../src/compiler/evaluator';
import { assert } from '../ts/test-runner';
import * as fs from 'fs';
import * as path from 'path';

const libCode = fs.readFileSync(
    path.join(__dirname, '../../sample/src/vba/LibString.bas'),
    'utf-8'
);

const ev = (() => {
    const tokens = new Lexer(libCode).tokenize();
    const ast = new Parser(tokens).parse();
    const ev = new Evaluator(console.log);
    ev.evaluate(ast);
    return ev;
})();

function r(name: string, ...args: any[]): any {
    return ev.callProcedure(name, args);
}

// --- IsNullOrEmpty ---
{
    assert.strictEqual(r('IsNullOrEmpty', ''),        -1, '空文字は empty');
    assert.strictEqual(r('IsNullOrEmpty', '   '),     -1, '空白のみは empty');
    assert.strictEqual(r('IsNullOrEmpty', 'hello'),    0, '通常文字列は非 empty');
    assert.strictEqual(r('IsNullOrEmpty', ' a '),      0, 'スペース込みでも内容あれば非 empty');
    console.log('[PASS] IsNullOrEmpty');
}

// --- IsAllDigits ---
{
    assert.strictEqual(r('IsAllDigits', '12345'),  -1, '全数字');
    assert.strictEqual(r('IsAllDigits', '0'),      -1, '0 単体');
    assert.strictEqual(r('IsAllDigits', '12a5'),    0, '文字混じり');
    assert.strictEqual(r('IsAllDigits', ''),        0, '空文字は false');
    assert.strictEqual(r('IsAllDigits', '1 2'),     0, 'スペース混じり');
    console.log('[PASS] IsAllDigits');
}

// --- IsAllAlpha ---
{
    assert.strictEqual(r('IsAllAlpha', 'abc'),   -1, '全アルファベット小文字');
    assert.strictEqual(r('IsAllAlpha', 'ABC'),   -1, '全アルファベット大文字');
    assert.strictEqual(r('IsAllAlpha', 'aBcD'),  -1, '混在でも全アルファベット');
    assert.strictEqual(r('IsAllAlpha', 'ab1'),    0, '数字混じり');
    assert.strictEqual(r('IsAllAlpha', ''),       0, '空文字は false');
    console.log('[PASS] IsAllAlpha');
}

// --- StartsWith / EndsWith ---
{
    assert.strictEqual(r('StartsWith', 'hello world', 'hello'),  -1, 'hello で始まる');
    assert.strictEqual(r('StartsWith', 'hello world', 'world'),   0, 'world では始まらない');
    assert.strictEqual(r('StartsWith', 'hello', ''),             -1, '空文字 prefix は常に true');
    assert.strictEqual(r('EndsWith',   'hello world', 'world'),  -1, 'world で終わる');
    assert.strictEqual(r('EndsWith',   'hello world', 'hello'),   0, 'hello では終わらない');
    assert.strictEqual(r('EndsWith',   'hello', ''),             -1, '空文字 suffix は常に true');
    console.log('[PASS] StartsWith / EndsWith');
}

// --- Contains ---
{
    assert.strictEqual(r('Contains', 'hello world', 'lo wo'),  -1, '部分文字列あり');
    assert.strictEqual(r('Contains', 'hello world', 'xyz'),     0, '部分文字列なし');
    assert.strictEqual(r('Contains', 'hello', ''),             -1, '空文字は常に含む');
    console.log('[PASS] Contains');
}

// --- CountOccurrences ---
{
    assert.strictEqual(r('CountOccurrences', 'abcabc', 'abc'),  2, '2回');
    assert.strictEqual(r('CountOccurrences', 'aaa', 'aa'),      1, '重複しない: 1回');
    assert.strictEqual(r('CountOccurrences', 'abc', 'xyz'),     0, '0回');
    assert.strictEqual(r('CountOccurrences', 'abc', ''),        0, '空文字 substr は 0');
    assert.strictEqual(r('CountOccurrences', 'a,b,,c', ','),    3, 'カンマ3個');
    console.log('[PASS] CountOccurrences');
}

// --- PadLeft / PadRight ---
{
    assert.strictEqual(r('PadLeft',  '42', 5, '0'), '00042', 'ゼロ左埋め');
    assert.strictEqual(r('PadLeft',  'abc', 3, '0'), 'abc',  'すでに十分な長さ');
    assert.strictEqual(r('PadRight', 'hi', 5, '-'), 'hi---', 'ハイフン右埋め');
    assert.strictEqual(r('PadRight', 'hello', 3, ' '), 'hello', 'すでに十分な長さ');
    console.log('[PASS] PadLeft / PadRight');
}

// --- ZeroPad ---
{
    assert.strictEqual(r('ZeroPad', 7, 3),   '007',   '3桁ゼロ埋め');
    assert.strictEqual(r('ZeroPad', 42, 5),  '00042', '5桁ゼロ埋め');
    assert.strictEqual(r('ZeroPad', 1234, 3), '1234', '桁数超えはそのまま');
    console.log('[PASS] ZeroPad');
}

// --- TrimAll ---
{
    assert.strictEqual(r('TrimAll', '  hello   world  '), 'hello world', '前後・中間の余分な空白を除去');
    assert.strictEqual(r('TrimAll', 'a  b  c'),           'a b c',      '連続スペースを1つに');
    assert.strictEqual(r('TrimAll', '   '),               '',           '空白のみは空文字');
    console.log('[PASS] TrimAll');
}

// --- Capitalize ---
{
    assert.strictEqual(r('Capitalize', 'hello'), 'Hello', '先頭大文字');
    assert.strictEqual(r('Capitalize', 'HELLO'), 'Hello', '全大文字→先頭のみ大文字');
    assert.strictEqual(r('Capitalize', 'a'),     'A',     '1文字');
    assert.strictEqual(r('Capitalize', ''),      '',      '空文字');
    console.log('[PASS] Capitalize');
}

// --- RepeatStr ---
{
    assert.strictEqual(r('RepeatStr', 'ab', 3), 'ababab', '3回繰り返し');
    assert.strictEqual(r('RepeatStr', 'x',  1), 'x',     '1回');
    assert.strictEqual(r('RepeatStr', 'x',  0), '',      '0回は空文字');
    console.log('[PASS] RepeatStr');
}

// --- ReverseString ---
{
    assert.strictEqual(r('ReverseString', 'hello'),  'olleh', '逆順');
    assert.strictEqual(r('ReverseString', 'a'),      'a',     '1文字');
    assert.strictEqual(r('ReverseString', ''),       '',      '空文字');
    assert.strictEqual(r('ReverseString', 'abba'),   'abba',  '回文');
    console.log('[PASS] ReverseString');
}

// --- Truncate ---
{
    assert.strictEqual(r('Truncate', 'hello world', 8, '...'), 'hello...', '切り詰め');
    assert.strictEqual(r('Truncate', 'hi', 10, '...'),         'hi',       '上限以下はそのまま');
    assert.strictEqual(r('Truncate', 'hello', 5, '...'),       'hello',    'ぴったりはそのまま');
    console.log('[PASS] Truncate');
}

// --- SubstringBefore / After ---
{
    assert.strictEqual(r('SubstringBefore', 'foo@bar.com', '@'), 'foo',     '@より前');
    assert.strictEqual(r('SubstringAfter',  'foo@bar.com', '@'), 'bar.com', '@より後');
    assert.strictEqual(r('SubstringBefore', 'foobar', '@'),      'foobar',  '区切りなし→全体');
    assert.strictEqual(r('SubstringAfter',  'foobar', '@'),      '',        '区切りなし→空文字');
    assert.strictEqual(r('SubstringBefore', 'a::b::c', '::'),    'a',       '最初の :: より前');
    assert.strictEqual(r('SubstringAfter',  'a::b::c', '::'),    'b::c',    '最初の :: より後');
    console.log('[PASS] SubstringBefore / SubstringAfter');
}

// --- SubstringBeforeLast / AfterLast ---
{
    assert.strictEqual(r('SubstringBeforeLast', 'a::b::c', '::'), 'a::b', '最後の :: より前');
    assert.strictEqual(r('SubstringAfterLast',  'a::b::c', '::'), 'c',    '最後の :: より後');
    assert.strictEqual(r('SubstringBeforeLast', 'foobar', '::'),  'foobar', '区切りなし→全体');
    assert.strictEqual(r('SubstringAfterLast',  'foobar', '::'),  '',       '区切りなし→空文字');
    console.log('[PASS] SubstringBeforeLast / SubstringAfterLast');
}

// --- WordCount ---
{
    assert.strictEqual(r('WordCount', 'hello world foo'), 3, '3単語');
    assert.strictEqual(r('WordCount', 'hello'),           1, '1単語');
    assert.strictEqual(r('WordCount', ''),                0, '空文字は0');
    assert.strictEqual(r('WordCount', '  '),              0, '空白のみは0');
    assert.strictEqual(r('WordCount', '  a  b  '),        2, '余分な空白は無視');
    console.log('[PASS] WordCount');
}

console.log('\n✅ LibString: 全テスト通過');
