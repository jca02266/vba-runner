import { Lexer } from '../../src/engine/lexer';
import { Parser } from '../../src/engine/parser';
import { findCallContext, SignatureHelpProvider } from '../../src/lsp/signature-help-provider';
import { assert } from '../../test-libs/test-runner';

function getSignatureHelp(src: string, line: number, character: number) {
    const tokens = new Lexer(src).tokenize();
    const ast = new Parser(tokens).parse();
    const provider = new SignatureHelpProvider();
    return provider.getSignatureHelp(ast.body, src, line, character);
}

// --- findCallContext の直接テスト ---

// 1. 基本: 関数名と activeParameter=0
{
    const ctx = findCallContext('Foo(', 4);
    assert.ok(ctx, 'Foo( → context あり');
    assert.strictEqual(ctx!.name, 'Foo', 'name = Foo');
    assert.strictEqual(ctx!.activeParameter, 0, 'activeParameter = 0');
    console.log('[PASS] findCallContext: 基本');
}

// 2. カンマ後は activeParameter=1
{
    const ctx = findCallContext('Foo(x, ', 7);
    assert.ok(ctx, 'Foo(x,  → context あり');
    assert.strictEqual(ctx!.activeParameter, 1, 'カンマ後は activeParameter=1');
    console.log('[PASS] findCallContext: カンマ後 activeParameter=1');
}

// 3. [回帰] 文字列内のカンマは activeParameter を増やさない (commit 595bcac)
// 修正前は "," の右→左スキャンで文字列内のカンマも数えていた
{
    const ctx = findCallContext('Foo("a, b", ', 12);
    assert.ok(ctx, 'Foo("a, b",  → context あり');
    assert.strictEqual(ctx!.name, 'Foo', 'name = Foo');
    assert.strictEqual(ctx!.activeParameter, 1,
        '文字列内の , は無視: activeParameter=1（2 ではない）');
    console.log('[PASS] findCallContext: 文字列内カンマを無視（false positive 回帰）');
}

// 4. [回帰] 文字列内の ( は深さカウントに影響しない (commit 595bcac)
{
    const ctx = findCallContext('Foo("x(y", ', 11);
    assert.ok(ctx, 'Foo("x(y",  → context あり');
    assert.strictEqual(ctx!.name, 'Foo', 'name = Foo（文字列内の ( で迷子にならない）');
    assert.strictEqual(ctx!.activeParameter, 1, 'activeParameter=1');
    console.log('[PASS] findCallContext: 文字列内 ( を無視（false positive 回帰）');
}

// 5. 閉じ括弧の後ろ（呼び出し対象外）→ null
{
    const ctx = findCallContext('x = Foo(a)', 10);
    assert.strictEqual(ctx, null, '閉じ括弧の後ろ → null');
    console.log('[PASS] findCallContext: 閉じ括弧後ろ → null');
}

// 6. ネスト呼び出し: 外側の引数位置を返す
{
    const ctx = findCallContext('Bar(Baz(x), ', 12);
    assert.ok(ctx, 'Bar(Baz(x),  → context あり');
    assert.strictEqual(ctx!.name, 'Bar', '外側 Bar が対象');
    assert.strictEqual(ctx!.activeParameter, 1, '2 番目の引数位置');
    console.log('[PASS] findCallContext: ネスト呼び出し外側');
}

// 7. 括弧なし → null
{
    const ctx = findCallContext('x = 1', 5);
    assert.strictEqual(ctx, null, '括弧なし → null');
    console.log('[PASS] findCallContext: 括弧なし → null');
}

// --- getSignatureHelp の統合テスト ---

// 8. ユーザー定義 Sub のシグネチャが返る
// 注: getSignatureHelp はカーソル行のみを参照するため、行内に開き括弧があれば十分
{
    const src = 'Sub Greet(name As String, count As Integer)\nEnd Sub\nSub Main()\n    Greet("hi", 1)\nEnd Sub';
    // 行 3: '    Greet("hi", 1)' — character 10 は開き ( の直後
    const result = getSignatureHelp(src, 3, 10);
    assert.ok(result, 'シグネチャヘルプが返る');
    const sig = result!.signature;
    assert.ok(sig, 'signature フィールドあり');
    assert.ok(sig.label.includes('Greet'), 'シグネチャに Greet が含まれる');
    assert.ok(sig.parameters && sig.parameters.length === 2, 'パラメーター数 2');
    console.log('[PASS] getSignatureHelp: ユーザー定義 Sub');
}

// 9. コメント行ではシグネチャヘルプを返さない
{
    const src = "Sub Foo(x As Long)\nEnd Sub\nSub Main()\n' Foo(\nEnd Sub";
    const result = getSignatureHelp(src, 3, 7);
    assert.strictEqual(result, null, 'コメント行 → null');
    console.log('[PASS] getSignatureHelp: コメント行 → null');
}

// 10. 未定義の関数名 → null
{
    const src = 'Sub Main()\n    x = Unknown("test")\nEnd Sub';
    // 行 1: '    x = Unknown("test")' — character 16 は Unknown の ( 直後
    const result = getSignatureHelp(src, 1, 16);
    assert.strictEqual(result, null, '未定義関数 → null');
    console.log('[PASS] getSignatureHelp: 未定義関数 → null');
}

// 11. VBA space-call スタイル: "MySub arg1" でシグネチャヘルプ
{
    const ctx = findCallContext('MySub arg1', 10);
    assert.ok(ctx, 'MySub arg1 → context あり');
    assert.strictEqual(ctx!.name, 'MySub', 'name = MySub');
    assert.strictEqual(ctx!.activeParameter, 0, 'activeParameter = 0');
    console.log('[PASS] findCallContext: space-call スタイル (第1引数)');
}

// 12. VBA space-call: カンマ後は activeParameter=1
{
    const ctx = findCallContext('MySub arg1, arg2', 16);
    assert.ok(ctx, 'MySub arg1, arg2 → context あり');
    assert.strictEqual(ctx!.name, 'MySub', 'name = MySub');
    assert.strictEqual(ctx!.activeParameter, 1, 'カンマ後は activeParameter=1');
    console.log('[PASS] findCallContext: space-call スタイル (第2引数)');
}

// 13. Call キーワード付き space-call
{
    const ctx = findCallContext('    Call MySub arg1', 19);
    assert.ok(ctx, 'Call MySub arg1 → context あり');
    assert.strictEqual(ctx!.name, 'MySub', 'name = MySub');
    assert.strictEqual(ctx!.activeParameter, 0, 'activeParameter = 0');
    console.log('[PASS] findCallContext: Call キーワード付き space-call');
}

// 14. space-call でも GetSignatureHelp が全体的に動作する
{
    const src = 'Sub Greet(name As String, count As Integer)\nEnd Sub\nSub Main()\n    Greet "hello", 1\nEnd Sub';
    // 行 3: '    Greet "hello", 1' — cursor at end (space-call style)
    const result = getSignatureHelp(src, 3, 19);
    assert.ok(result, 'Greet "hello", 1 → シグネチャヘルプあり');
    assert.strictEqual(result!.activeParameter, 1, '2番目の引数でアクティブパラメーター=1');
    console.log('[PASS] getSignatureHelp: space-call スタイル');
}

console.log('\n✅ lsp-signature-help: 全テスト通過');
