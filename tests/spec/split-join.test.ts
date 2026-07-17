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

    // 仕様バグ修正: 下限付き固定配列（Dim a(1 To n)）は物理ストレージに
    // LBound 分の隠し要素（添字0）を含むため、Join がそのまま arr.join() すると
    // 先頭に余分な空要素が混入していた。vbaBase を見てスキップするよう修正。
    {
        const code = `
            Dim parts(1 To 3) As String
            parts(1) = "a"
            parts(2) = "b"
            parts(3) = "c"
            Dim result As String
            result = Join(parts, "; ")
        `;
        const ev2 = evalVBA(code);
        assert.strictEqual(ev2.env.get('result'), 'a; b; c', 'Dim a(1 To n) の Join に余分な空要素が混入しない');
    }

    console.log('[PASS] Join');
}

// --- Split と Join の往復 ---
{
    assert.strictEqual(ev(`Join(Split("a,b,c", ","), "-")`), 'a-b-c', 'Split→Join 区切り変更');
    assert.strictEqual(ev(`Join(Split("a,b,c", ","), ",")`), 'a,b,c', 'Split→Join 同じ区切り');
    console.log('[PASS] Split と Join の往復');
}

// --- Bug #25-5: Split の limit 引数 ---
{
    const runFunc = (code: string, name: string, args: any[] = []) =>
        evalVBASingle(code).callProcedure(name, args);

    const r2 = runFunc(`
    Function T()
        Dim a() As String
        a = Split("A B C D", " ", 2)
        T = a(0) & "|" & a(1)
    End Function`, 'T');
    assert.strictEqual(r2, 'A|B C D', 'Split limit=2: 最後の要素に残り全体が入る');

    const r1 = runFunc(`
    Function T()
        Dim a() As String
        a = Split("A B C", " ", 1)
        T = a(0)
    End Function`, 'T');
    assert.strictEqual(r1, 'A B C', 'Split limit=1: 全体が1要素');

    const r0 = runFunc(`
    Function T()
        Dim a() As String
        a = Split("A B C", " ", -1)
        T = UBound(a)
    End Function`, 'T');
    assert.strictEqual(r0, 2, 'Split limit=-1: 全分割（デフォルト同様）');

    console.log('[PASS] Bug #25-5: Split limit 引数');
}

// --- Bug CK: 空 Expression / 空 Delimiter / Compare 無視 / Null 引数 (§6.1.2.11.1.35) ---
{
    const errOf = (expr: string) => {
        try { ev(expr); return 0; } catch (e: any) { return e?.number ?? -1; }
    };

    // 仕様: Expression が zero-length string → 要素なしの空配列
    assert.strictEqual(ev('UBound(Split("", ","))'), -1, 'Split("", ","): 空配列 (UBound = -1)');
    assert.strictEqual(ev('LBound(Split("", ","))'), 0, 'Split("", ","): LBound = 0');

    // 仕様: Delimiter が zero-length string → expression 全体を含む1要素配列
    assert.strictEqual(ev('UBound(Split("abc", ""))'), 0, 'Split("abc", ""): 1要素');
    assert.strictEqual(ev('Join(Split("abc", ""), "|")'), 'abc', 'Split("abc", ""): 全体が1要素');

    // 仕様: Compare = vbTextCompare → 区切り文字を大文字小文字無視で照合
    assert.strictEqual(ev('Join(Split("aXbXc", "x", -1, 1), "|")'), 'a|b|c', 'Split vbTextCompare: 大文字区切りにマッチ');
    assert.strictEqual(ev('Join(Split("aXbxc", "X", 2, 1), "|")'), 'a|bxc', 'Split vbTextCompare + limit: 残余は元の文字列を保持');
    // 省略時は binary（既定）
    assert.strictEqual(ev('UBound(Split("aXbXc", "x"))'), 0, 'Split binary: 大文字区切りにマッチしない');

    // Delimiter / Limit / Compare に Null → error 94
    assert.strictEqual(errOf('Split("a b", Null)'), 94, 'Split: Delimiter=Null -> error 94');
    assert.strictEqual(errOf('Split("a b", " ", Null)'), 94, 'Split: Limit=Null -> error 94');
    assert.strictEqual(errOf('Split("a b", " ", -1, Null)'), 94, 'Split: Compare=Null -> error 94');

    console.log('[PASS] Bug CK: Split 空文字列・Compare・Null 引数');
}

// Join: 非配列引数 → Type Mismatch 13 (spec: SourceArray must be array)
{
    const errOf = (expr: string) => { try { ev(expr); return null; } catch(e: any) { return e.number ?? null; } };
    assert.strictEqual(errOf('Join("notanarray", ",")'), 13, 'Join: 非配列引数 → Type Mismatch 13');
    console.log('[PASS] Join: 非配列引数はエラー');
}

// Join: 配列に Null 要素が含まれる → Type Mismatch 13（JS内部TypeErrorではなくVBAエラー）
{
    const errOf = (expr: string) => { try { ev(expr); return null; } catch(e: any) { return e.number ?? -1; } };
    assert.strictEqual(errOf('Join(Array(1, Null, 3), ",")'), 13, 'Join: Null要素 → Type Mismatch 13');
    console.log('[PASS] Join: Null要素はType Mismatch 13');
}

console.log('\n✅ Split / Join: 全テスト通過');
