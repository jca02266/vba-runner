/**
 * 組み込み文字列関数のテスト (§6.1.2.11)
 *
 * 対象機能（TODO.md より）:
 *   - Asc / AscW                  (§6.1.2.11.1.1)
 *   - Chr / Chr$ / ChrW           (§6.1.2.11.1.4/6)
 *   - InStr / InStrB              (§6.1.2.11.1.14)
 *   - LCase / LCase$              (§6.1.2.11.1.17/18)
 *   - Left / Left$                (§6.1.2.11.1.19/20)
 *   - Len                         (§6.1.2.11.1.22)
 *   - Mid / Mid$                  (§6.1.2.11.1.25/26)
 *   - Right / Right$              (§6.1.2.11.1.30/31)
 *   - Space / Space$              (§6.1.2.11.1.33/34)
 *   - String / String$            (§6.1.2.11.1.38/39)
 *   - UCase / UCase$              (§6.1.2.11.1.41/42)
 */
import { Lexer } from '../../src/engine/lexer';
import { Parser } from '../../src/engine/parser';
import { Evaluator } from '../../src/engine/evaluator';
import { assert } from '../../test-libs/test-runner';

function evalVBA(code: string): any {
    const tokens = new Lexer(code).tokenize();
    const ast = new Parser(tokens).parse();
    const ev = new Evaluator(console.log);
    ev.evaluate(ast);
    return ev;
}

function ev(expr: string): any {
    return evalVBA('').evalExpression(expr);
}

// --- Asc / AscW ---
{
    assert.strictEqual(ev(`Asc("A")`), 65, 'Asc("A")');
    assert.strictEqual(ev(`Asc("a")`), 97, 'Asc("a")');
    assert.strictEqual(ev(`Asc("0")`), 48, 'Asc("0")');
    assert.strictEqual(ev(`Asc("ABC")`), 65, 'Asc は先頭文字のコード');
    assert.strictEqual(ev(`AscW("A")`), 65, 'AscW("A")');
    assert.strictEqual(ev(`AscW("あ")`), 12354, 'AscW("あ") = U+3042');
    console.log('[PASS] Asc / AscW');
}

// --- Chr / Chr$ / ChrW ---
{
    assert.strictEqual(ev(`Chr(65)`), 'A', 'Chr(65) = "A"');
    assert.strictEqual(ev(`Chr(97)`), 'a', 'Chr(97) = "a"');
    assert.strictEqual(ev(`Chr$(65)`), 'A', 'Chr$(65) = "A"');
    assert.strictEqual(ev(`ChrW(12354)`), 'あ', 'ChrW(12354) = "あ"');
    assert.strictEqual(ev(`Chr(Asc("Z"))`), 'Z', 'Chr(Asc("Z")) = "Z" (往復)');
    console.log('[PASS] Chr / Chr$ / ChrW');
}

// --- InStr / InStrB ---
{
    // 2引数版: InStr(string1, string2)
    assert.strictEqual(ev(`InStr("Hello World", "World")`), 7, 'InStr("Hello World", "World") = 7');
    assert.strictEqual(ev(`InStr("Hello", "xyz")`), 0, '見つからない場合は 0');
    assert.strictEqual(ev(`InStr("Hello", "")`), 1, '空文字列の検索は 1');
    // 3引数版: InStr(start, string1, string2)
    assert.strictEqual(ev(`InStr(1, "abcabc", "b")`), 2, 'InStr(1, "abcabc", "b") = 2');
    assert.strictEqual(ev(`InStr(3, "abcabc", "b")`), 5, 'InStr(3, "abcabc", "b") = 5');
    // InStrB
    assert.strictEqual(ev(`InStrB("Hello", "ll")`), 5, 'InStrB はバイト単位（ASCII では 1 文字=2 バイト）');
    console.log('[PASS] InStr / InStrB');
}

// --- LCase / LCase$ ---
{
    assert.strictEqual(ev(`LCase("HELLO")`), 'hello', 'LCase("HELLO")');
    assert.strictEqual(ev(`LCase("Hello World")`), 'hello world', 'LCase 混在');
    assert.strictEqual(ev(`LCase("abc")`), 'abc', 'LCase すでに小文字');
    assert.strictEqual(ev(`LCase("")`), '', 'LCase 空文字列');
    assert.strictEqual(ev(`LCase$("HELLO")`), 'hello', 'LCase$ は LCase と同じ');
    console.log('[PASS] LCase / LCase$');
}

// --- Left / Left$ ---
{
    assert.strictEqual(ev(`Left("Hello World", 5)`), 'Hello', 'Left("Hello World", 5)');
    assert.strictEqual(ev(`Left("ABC", 1)`), 'A', 'Left("ABC", 1)');
    assert.strictEqual(ev(`Left("ABC", 0)`), '', 'Left 長さ 0');
    assert.strictEqual(ev(`Left("ABC", 10)`), 'ABC', 'Left 長さ > 文字列長');
    assert.strictEqual(ev(`Left$("Hello", 3)`), 'Hel', 'Left$ は Left と同じ');
    console.log('[PASS] Left / Left$');
}

// --- Len ---
{
    assert.strictEqual(ev(`Len("Hello")`), 5, 'Len("Hello") = 5');
    assert.strictEqual(ev(`Len("")`), 0, 'Len("") = 0');
    assert.strictEqual(ev(`Len("あいう")`), 3, 'Len は文字数（Unicode 文字）');
    console.log('[PASS] Len');
}

// --- Mid / Mid$ ---
{
    // 3引数版: Mid(string, start, length)
    assert.strictEqual(ev(`Mid("Hello World", 1, 5)`), 'Hello', 'Mid(s, 1, 5)');
    assert.strictEqual(ev(`Mid("Hello World", 7, 5)`), 'World', 'Mid(s, 7, 5)');
    assert.strictEqual(ev(`Mid("ABCDE", 2, 2)`), 'BC', 'Mid(s, 2, 2)');
    // 2引数版: Mid(string, start) - start 以降全部
    assert.strictEqual(ev(`Mid("Hello World", 7)`), 'World', 'Mid(s, 7) start 以降全部');
    // Mid$
    assert.strictEqual(ev(`Mid$("Hello", 2, 3)`), 'ell', 'Mid$ は Mid と同じ');
    console.log('[PASS] Mid / Mid$');
}

// --- Right / Right$ ---
{
    assert.strictEqual(ev(`Right("Hello World", 5)`), 'World', 'Right("Hello World", 5)');
    assert.strictEqual(ev(`Right("ABC", 1)`), 'C', 'Right("ABC", 1)');
    assert.strictEqual(ev(`Right("ABC", 0)`), '', 'Right 長さ 0');
    assert.strictEqual(ev(`Right("ABC", 10)`), 'ABC', 'Right 長さ > 文字列長');
    assert.strictEqual(ev(`Right$("Hello", 3)`), 'llo', 'Right$ は Right と同じ');
    console.log('[PASS] Right / Right$');
}

// --- Space / Space$ ---
{
    assert.strictEqual(ev(`Space(3)`), '   ', 'Space(3) = "   "');
    assert.strictEqual(ev(`Space(0)`), '', 'Space(0) = ""');
    assert.strictEqual(ev(`Space(1)`), ' ', 'Space(1) = " "');
    assert.strictEqual(ev(`Space$(5)`), '     ', 'Space$ は Space と同じ');
    console.log('[PASS] Space / Space$');
}

// --- String / String$ ---
{
    // 文字コードを指定
    assert.strictEqual(ev(`String(5, 65)`), 'AAAAA', 'String(5, 65) = "AAAAA"');
    // 文字を指定（先頭文字のみ使用）
    assert.strictEqual(ev(`String(5, "*")`), '*****', 'String(5, "*")');
    assert.strictEqual(ev(`String(3, "ABC")`), 'AAA', 'String 文字列の先頭のみ使用');
    assert.strictEqual(ev(`String(0, "*")`), '', 'String(0, "*") = ""');
    assert.strictEqual(ev(`String$(4, "X")`), 'XXXX', 'String$ は String と同じ');
    console.log('[PASS] String / String$');
}

// --- UCase / UCase$ ---
{
    assert.strictEqual(ev(`UCase("hello")`), 'HELLO', 'UCase("hello")');
    assert.strictEqual(ev(`UCase("Hello World")`), 'HELLO WORLD', 'UCase 混在');
    assert.strictEqual(ev(`UCase("ABC")`), 'ABC', 'UCase すでに大文字');
    assert.strictEqual(ev(`UCase("")`), '', 'UCase 空文字列');
    assert.strictEqual(ev(`UCase$("hello")`), 'HELLO', 'UCase$ は UCase と同じ');
    console.log('[PASS] UCase / UCase$');
}

// vbBack / vbFormFeed 定数
{
    assert.strictEqual(ev('Asc(vbBack)'), 8, 'vbBack = Chr(8)');
    assert.strictEqual(ev('Asc(vbFormFeed)'), 12, 'vbFormFeed = Chr(12)');
    console.log('[PASS] vbBack / vbFormFeed');
}

console.log('\n✅ 組み込み文字列関数: 全テスト通過');
