/**
 * Split / Join 関数のテスト (§6.1.2.11.1.16 / §6.1.2.11.1.35)
 *
 * VBA 標準の文字列⇔配列変換関数。
 */
import { evalVBASingle, assert } from '../../test-libs/test-runner';

function evalVBA(code: string): any {
    return evalVBASingle(code);
}

function ev(expr: string): any {
    return evalVBA('').evalExpression(expr);
}

console.log('--- Starting Split / Join Tests ---');

// --- Split ---
{
    // 基本: 区切り文字で分割
    const parts1 = ev(`Split("a,b,c", ",")`);
    assert.strictEqual(Array.isArray(parts1), true, 'Split は配列を返す');
    assert.strictEqual(parts1.length, 3, 'Split("a,b,c", ",") は 3 要素');
    assert.strictEqual(parts1[0], 'a', 'parts[0] = "a"');
    assert.strictEqual(parts1[1], 'b', 'parts[1] = "b"');
    assert.strictEqual(parts1[2], 'c', 'parts[2] = "c"');

    // 空白区切り
    const parts2 = ev(`Split("hello world foo", " ")`);
    assert.strictEqual(parts2.length, 3, 'スペース区切り 3 要素');
    assert.strictEqual(parts2[0], 'hello', 'parts2[0]');
    assert.strictEqual(parts2[2], 'foo', 'parts2[2]');

    // 複数文字の区切り
    const parts3 = ev(`Split("a::b::c", "::")`);
    assert.strictEqual(parts3.length, 3, '複数文字区切り 3 要素');
    assert.strictEqual(parts3[1], 'b', 'parts3[1] = "b"');

    // 区切り文字なし → 1 要素
    const parts4 = ev(`Split("abc", ",")`);
    assert.strictEqual(parts4.length, 1, '区切り文字なし → 1 要素');
    assert.strictEqual(parts4[0], 'abc', 'parts4[0] = "abc"');

    // 連続する区切り文字 → 空要素
    const parts5 = ev(`Split("a,,b", ",")`);
    assert.strictEqual(parts5.length, 3, '連続区切り → 3 要素');
    assert.strictEqual(parts5[1], '', '空要素');

    console.log('[PASS] Split');
}

// --- Join ---
{
    // 配列を区切り文字で連結
    assert.strictEqual(ev(`Join(Array("a", "b", "c"), ",")`), 'a,b,c', 'Join("a,b,c", ",")');
    assert.strictEqual(ev(`Join(Array("hello", "world"), " ")`), 'hello world', 'スペース連結');
    assert.strictEqual(ev(`Join(Array("a", "b", "c"), "::")`), 'a::b::c', '複数文字区切り');

    // 数値配列
    assert.strictEqual(ev(`Join(Array(1, 2, 3), "-")`), '1-2-3', '数値配列の連結');

    // 1 要素
    assert.strictEqual(ev(`Join(Array("only"), ",")`), 'only', '1 要素');

    // 区切り省略時はスペース
    assert.strictEqual(ev(`Join(Array("a", "b", "c"))`), 'a b c', '区切り省略 = スペース');

    console.log('[PASS] Join');
}

// --- Split と Join の往復 ---
{
    assert.strictEqual(ev(`Join(Split("a,b,c", ","), "-")`), 'a-b-c', 'Split→Join 区切り変更');
    assert.strictEqual(ev(`Join(Split("a,b,c", ","), ",")`), 'a,b,c', 'Split→Join 同じ区切り');
    console.log('[PASS] Split と Join の往復');
}

console.log('\n✅ Split / Join: 全テスト通過');
