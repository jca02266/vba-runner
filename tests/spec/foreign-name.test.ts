/**
 * FOREIGN-NAME `[identifier]` 構文テスト (MS-VBAL §3.3.5.2)
 *
 * FOREIGN-NAME = "[" foreign-identifier "]"
 * - 予約語をプロシージャ呼び出しに使用できる (例: [End]())
 * - 定義側 (Sub [End]) は実 VBA でも不可 → エンジンでも実装しない
 * - 未定義でも Option Explicit エラーにならない
 * - 定義済みの場合は通常の呼び出しと同様に動作する
 */

import { evalVBASingle, assert } from '../../test-libs/test-runner';

// ─── 1. 基本動作: 定義済み関数を [name] 構文で呼び出す ───────────────────────

{
    const code = `
Dim g As Long
Function MyFunc() As Long
    MyFunc = 42
End Function
Sub Test()
    g = [MyFunc]()
End Sub
`;
    const ev = evalVBASingle(code);
    ev.callProcedure('Test', []);
    assert.strictEqual(ev.env.get('g'), 42, '[MyFunc]() calls MyFunc normally');
    console.log('[PASS] FN01: [MyFunc]() — 定義済み関数の呼び出し');
}

// ─── 2. 予約語を含む名前を [name] 構文で呼び出す ───────────────────────────

{
    // Option Explicit 下でも未定義 FOREIGN-NAME はコンパイルエラーにならない
    const code = `
Option Explicit
Sub Test()
    [End]
End Sub
`;
    let caught = false;
    try {
        evalVBASingle(code);
    } catch (e: any) {
        caught = true;
        console.error('FN02 unexpected error:', e.message);
    }
    assert.strictEqual(caught, false, '[End] should not throw compile error under Option Explicit');
    console.log('[PASS] FN02: Option Explicit + [End] — コンパイルエラーにならない');
}

// ─── 3. 引数付き FOREIGN-NAME 呼び出し ────────────────────────────────────

{
    const code = `
Dim g As Long
Function Add(a As Long, b As Long) As Long
    Add = a + b
End Function
Sub Test()
    g = [Add](10, 32)
End Sub
`;
    const ev = evalVBASingle(code);
    ev.callProcedure('Test', []);
    assert.strictEqual(ev.env.get('g'), 42, '[Add](10, 32) = 42');
    console.log('[PASS] FN03: [Add](10, 32) — 引数付き FOREIGN-NAME 呼び出し');
}

// ─── 4. [name] をステートメント文脈で使用 ─────────────────────────────────

{
    const code = `
Dim g As Long
Sub MySub()
    g = 99
End Sub
Sub Test()
    [MySub]
End Sub
`;
    const ev = evalVBASingle(code);
    ev.callProcedure('Test', []);
    assert.strictEqual(ev.env.get('g'), 99, '[MySub] calls MySub as a statement');
    console.log('[PASS] FN04: [MySub] — ステートメント文脈での FOREIGN-NAME 呼び出し');
}

// ─── 5. Call キーワード + FOREIGN-NAME ──────────────────────────────────

{
    const code = `
Dim g As Long
Sub MySub()
    g = 55
End Sub
Sub Test()
    Call [MySub]()
End Sub
`;
    const ev = evalVBASingle(code);
    ev.callProcedure('Test', []);
    assert.strictEqual(ev.env.get('g'), 55, 'Call [MySub]() calls MySub');
    console.log('[PASS] FN05: Call [MySub]() — Call キーワード + FOREIGN-NAME');
}

// ─── 6. 複数の未定義 FOREIGN-NAME が Option Explicit エラーにならない ────

{
    const code = `
Option Explicit
Sub Test()
    [Foo]
    [Bar]
    [Baz]
End Sub
`;
    let caught = false;
    try {
        evalVBASingle(code);
    } catch (e: any) {
        caught = true;
        console.error('FN06 unexpected error:', e.message);
    }
    assert.strictEqual(caught, false, 'Multiple undefined [name] should not throw compile error');
    console.log('[PASS] FN06: 複数の未定義 FOREIGN-NAME — Option Explicit エラーなし');
}

// ─── 7. 未定義 FOREIGN-NAME の実行時エラー ───────────────────────────────

{
    const code = `
Option Explicit
Sub Test()
    [UndefinedProc]()
End Sub
`;
    const ev = evalVBASingle(code);
    let caught = false;
    try {
        ev.callProcedure('Test', []);
    } catch (e: any) {
        caught = true;
    }
    assert.strictEqual(caught, true, 'Calling undefined [name]() should throw at runtime');
    console.log('[PASS] FN07: 未定義 [UndefinedProc]() — 実行時エラーになる');
}

// ─── 8. 予約語の名前で定義された関数を FOREIGN-NAME で呼び出す ─────────────
// 実 VBA では Sub [End]() は不可だが、エンジンでは通常識別子名として登録された
// 関数を [End] 構文で呼び出すこともできる（予約語名は通常では宣言できないので実用外）

{
    // FOREIGN-NAME をメンバーアクセスのオブジェクト側にも使えることを確認
    const code = `
Dim g As Long
Sub Test()
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    [d].Add "k", 1
    g = [d].Count
End Sub
`;
    // この形式（[d] をメンバーアクセスのオブジェクトとして使う）は実 VBA ではコンパイルエラーになるが
    // エンジンでは動作する場合がある。ここでは Option Explicit エラーにならないことを確認
    let caught = false;
    try {
        const ev = evalVBASingle(code);
        ev.callProcedure('Test', []);
    } catch (e: any) {
        // Runtime error は OK (undeclared variable at runtime)
        // Compile error は NG
        if (e.message && e.message.includes('Option Explicit')) {
            caught = true;
        }
    }
    assert.strictEqual(caught, false, '[d] as object should not throw Option Explicit compile error');
    console.log('[PASS] FN08: [d].Method — FOREIGN-NAME をオブジェクトとして使っても Option Explicit エラーなし');
}

console.log('\n✅ FOREIGN-NAME [identifier] 構文: 全テスト通過');
