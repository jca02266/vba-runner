/**
 * 組み込み文字列関数のテスト (§6.1.2.11)
 *
 * 対象機能（TODO_SPEC.md より）:
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
import { evalVBASingle, assert } from '../../test-libs/test-runner';
import { vbaNull } from '../../src/engine/evaluator';

function evalVBA(code: string): any {
    return evalVBASingle(code);
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
    assert.strictEqual(ev(`ChrW$(169)`), '©', 'ChrW$(169) = "©"');
    assert.strictEqual(ev(`Chr(Asc("Z"))`), 'Z', 'Chr(Asc("Z")) = "Z" (往復)');
    console.log('[PASS] Chr / Chr$ / ChrW / ChrW$');
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

// --- Bug #25-1〜3: LenB / AscB / ChrB (UTF-16LE モデル) ---
{
    assert.strictEqual(ev('LenB("ABC")'), 6, 'LenB("ABC") = 6 (3文字×2バイト)');
    assert.strictEqual(ev('LenB("")'), 0, 'LenB("") = 0');
    assert.strictEqual(ev('LenB("A")'), 2, 'LenB("A") = 2');
    assert.strictEqual(ev('AscB("A")'), 65, 'AscB("A") = 65');
    assert.strictEqual(ev('AscB("Z")'), 90, 'AscB("Z") = 90');
    assert.strictEqual(ev('ChrB(65)'), 'A', 'ChrB(65) = "A"');
    assert.strictEqual(ev('ChrB(90)'), 'Z', 'ChrB(90) = "Z"');
    assert.strictEqual(ev('AscB(ChrB(65))'), 65, 'AscB(ChrB(65)) = 65 (往復)');
    console.log('[PASS] Bug #25-1〜3: LenB / AscB / ChrB');
}

// --- Bug A: Replace — start / count / compare 引数 ---
{
    // start: 返り値は start 位置以降のみ（prefix は含まれない）
    assert.strictEqual(ev('Replace("Hello Hello Hello", "Hello", "Hi", 7)'), 'Hi Hi', 'Replace start=7: 7文字目以降で全置換');
    assert.strictEqual(ev('Replace("Hello Hello Hello", "Hello", "Hi", 7, 1)'), 'Hi Hello', 'Replace start=7 count=1: 1件のみ置換');
    // count
    assert.strictEqual(ev('Replace("Hello Hello Hello", "Hello", "Hi", 1, 1)'), 'Hi Hello Hello', 'Replace count=1: 最初の1件のみ');
    assert.strictEqual(ev('Replace("Hello Hello Hello", "Hello", "Hi", 1, 2)'), 'Hi Hi Hello', 'Replace count=2: 2件のみ');
    assert.strictEqual(ev('Replace("Hello Hello Hello", "Hello", "Hi", 1, 0)'), 'Hello Hello Hello', 'Replace count=0: 置換なし');
    // compare (vbTextCompare=1)
    assert.strictEqual(ev('Replace("Hello hello", "HELLO", "Hi", 1, -1, 1)'), 'Hi Hi', 'Replace vbTextCompare: 大文字小文字無視');
    assert.strictEqual(ev('Replace("Hello hello", "HELLO", "Hi", 1, -1, 0)'), 'Hello hello', 'Replace vbBinaryCompare: 大文字小文字区別');
    // find が空文字列 → working をそのまま返す
    assert.strictEqual(ev('Replace("Hello", "", "x", 1)'), 'Hello', 'Replace find="" → そのまま返す');
    console.log('[PASS] Bug A: Replace start/count/compare');
}

// --- Bug N: Left/Right/Mid の Null 伝播 ---
{
    const checkIsNull = (expr: string) => {
        const env2 = evalVBA(`\nFunction T()\n    If IsNull(${expr}) Then\n        T = 1\n    Else\n        T = 0\n    End If\nEnd Function\n`);
        return env2.callProcedure('T', []);
    };
    assert.strictEqual(checkIsNull('Left(Null, 2)'), 1, 'Left(Null) → IsNull=True');
    assert.strictEqual(checkIsNull('Right(Null, 2)'), 1, 'Right(Null) → IsNull=True');
    assert.strictEqual(checkIsNull('Mid(Null, 1)'), 1, 'Mid(Null) → IsNull=True');
    assert.strictEqual(ev('Left("hello", 2)'), 'he', 'Left 通常動作');
    assert.strictEqual(ev('Right("hello", 2)'), 'lo', 'Right 通常動作');
    assert.strictEqual(ev('Mid("hello", 2, 3)'), 'ell', 'Mid 通常動作');
    console.log('[PASS] Bug N: Left/Right/Mid Null 伝播');
}

// --- Bug H: Asc("") は Error 5 ---
{
    assert.throwsMatch(() => ev('Asc("")'), /error '5'/, 'Asc("") → Error 5');
    assert.strictEqual(ev('Asc("A")'), 65, 'Asc 通常動作');
    console.log('[PASS] Bug H: Asc("") → Error 5');
}

// --- Bug I: Left/Right/Mid の負値引数は Error 5 ---
{
    assert.throwsMatch(() => ev('Left("abc", -1)'), /error '5'/, 'Left 負長さ → Error 5');
    assert.throwsMatch(() => ev('Right("abc", -1)'), /error '5'/, 'Right 負長さ → Error 5');
    assert.throwsMatch(() => ev('Mid("abc", 0)'), /error '5'/, 'Mid start=0 → Error 5');
    assert.throwsMatch(() => ev('Mid("abc", 1, -1)'), /error '5'/, 'Mid 負長さ → Error 5');
    assert.strictEqual(ev('Left("abc", 0)'), '', 'Left 長さ0 → ""');
    assert.strictEqual(ev('Mid("abc", 1, 0)'), '', 'Mid 長さ0 → ""');
    console.log('[PASS] Bug I: Left/Right/Mid 負値引数 → Error 5');
}

// --- Bug J: Space/String の負値引数は Error 5 ---
{
    assert.throwsMatch(() => ev('Space(-1)'), /error '5'/, 'Space(-1) → Error 5');
    assert.throwsMatch(() => ev('String(-1, "x")'), /error '5'/, 'String(-1) → Error 5');
    assert.strictEqual(ev('Space(0)'), '', 'Space(0) = ""');
    assert.strictEqual(ev('String(0, "x")'), '', 'String(0) = ""');
    console.log('[PASS] Bug J: Space/String 負値引数 → Error 5');
}

// --- Bug K: InStr の start < 1 は Error 5 ---
{
    assert.throwsMatch(() => ev('InStr(0, "abc", "b")'), /error '5'/, 'InStr(0) → Error 5');
    assert.throwsMatch(() => ev('InStr(-1, "abc", "b")'), /error '5'/, 'InStr(-1) → Error 5');
    assert.strictEqual(ev('InStr(1, "abc", "b")'), 2, 'InStr(1) 通常動作');
    assert.strictEqual(ev('InStr("abc", "b")'), 2, 'InStr 2引数版 通常動作');
    console.log('[PASS] Bug K: InStr start<1 → Error 5');
}

// --- Bug M: Chr の引数範囲は 0-255、ChrW は 0-65535 ---
{
    assert.strictEqual(ev('Chr(65)'), 'A', 'Chr(65) = "A"');
    assert.strictEqual(ev('Chr(255)'), 'ÿ', 'Chr(255) = "ÿ"');
    assert.throwsMatch(() => ev('Chr(256)'), /error '5'/, 'Chr(256) → Error 5');
    assert.throwsMatch(() => ev('Chr(-1)'), /error '5'/, 'Chr(-1) → Error 5');
    assert.strictEqual(ev('ChrW(256)'), 'Ā', 'ChrW(256) = "Ā"');
    assert.strictEqual(ev('ChrW(12354)'), 'あ', 'ChrW(12354) = "あ"');
    assert.throwsMatch(() => ev('ChrW(65536)'), /error '5'/, 'ChrW(65536) → Error 5');
    console.log('[PASS] Bug M: Chr(>255) / ChrW(>65535) → Error 5');
}

// --- Bug S: LenB(Null) が Null を返さずゴミ値を返す ---
{
    assert.strictEqual(ev('LenB(Null)') === vbaNull, true, 'LenB(Null) = Null');
    assert.strictEqual(ev('Len(Null)') === vbaNull, true, 'Len(Null) = Null');
    assert.strictEqual(ev('LenB("hello")'), 10, 'LenB("hello") = 10');
    console.log('[PASS] Bug S: LenB(Null) = Null');
}

// --- Bug AJ: ChrB(Null) が TypeError でクラッシュする ---
{
    assert.strictEqual(ev('ChrB(Null)') === vbaNull, true, 'ChrB(Null) = Null');
    assert.strictEqual(ev('ChrB(65)'), 'A', 'ChrB(65) = "A"');
    console.log('[PASS] Bug AJ: ChrB(Null) = Null');
}

// --- Bug AP/AQ/AR: Left/Right/Mid の len/start 引数が Null のとき Null を返す ---
{
    assert.strictEqual(ev('Left("abc", Null)') === vbaNull, true, 'Left(str, Null) = Null');
    assert.strictEqual(ev('Right("abc", Null)') === vbaNull, true, 'Right(str, Null) = Null');
    assert.strictEqual(ev('Mid("abc", 1, Null)') === vbaNull, true, 'Mid(str, start, Null) = Null');
    assert.strictEqual(ev('Mid("abc", Null, 1)') === vbaNull, true, 'Mid(str, Null, len) = Null');
    console.log('[PASS] Bug AP/AQ/AR: Left/Right/Mid の Null len/start 伝播');
}

// --- Bug AU/AV/AW: LeftB/RightB/MidB の len/start 引数が Null のとき Null を返す ---
{
    assert.strictEqual(ev('LeftB("abc", Null)') === vbaNull, true, 'LeftB(str, Null) = Null');
    assert.strictEqual(ev('RightB("abc", Null)') === vbaNull, true, 'RightB(str, Null) = Null');
    assert.strictEqual(ev('MidB("abc", 1, Null)') === vbaNull, true, 'MidB(str, start, Null) = Null');
    assert.strictEqual(ev('MidB("abc", Null, 2)') === vbaNull, true, 'MidB(str, Null, len) = Null');
    console.log('[PASS] Bug AU/AV/AW: LeftB/RightB/MidB の Null len/start 伝播');
}

// --- Bug AM/AN: Split(Null) / Join(arr, Null) の Null 伝播 ---
{
    assert.strictEqual(ev('Split(Null, ",")') === vbaNull, true, 'Split(Null, ",") = Null');
    assert.strictEqual(ev('Join(Array("a", "b"), Null)') === vbaNull, true, 'Join(arr, Null) = Null');
    console.log('[PASS] Bug AM/AN: Split/Join の Null 伝播');
}

console.log('\n✅ 組み込み文字列関数: 全テスト通過');
