/**
 * Bugs 32-A〜32-D（評価 #32・差し込みテンプレートエンジンドメイン）のリグレッションテスト。
 *
 * - 32-A: `Tab(n)` が n+1 桁目に出力（オフバイワン）。Print # / Debug.Print 共通
 * - 32-B: `Write #1, #2024/03/15#` でレクサーが `#1, #` を日付リテラルと誤認しパースエラー
 * - 32-C: `Input #` が引用符内のカンマで分割して Error 13
 * - 32-D: `Write #` の日付が universal format（#yyyy-mm-dd#）でなく、改行が LF（Print # は CRLF）
 */
import { evalVBASingle, assert } from '../../test-libs/test-runner';

// --- Bug 32-A: Tab(n) は次の出力を n 桁目（1 始まり）に置く ---
{
    const out: string[] = [];
    const ev = evalVBASingle(`
Sub T()
    Debug.Print "X"; Tab(10); "Y"
    Debug.Print "X"; Spc(8); "Y"
    Debug.Print "AB", "C"
End Sub`, { onPrint: (s: string) => out.push(s) });
    ev.callProcedure('T', []);
    assert.strictEqual(out[0].indexOf('Y') + 1, 10, 'Bug 32-A: Tab(10) で Y が 10 桁目');
    assert.strictEqual(out[1].indexOf('Y') + 1, 10, 'Spc(8) で X の後に空白 8 → Y は 10 桁目');
    assert.strictEqual(out[2].indexOf('C') + 1, 15, 'カンマゾーンは 14 桁区切り（C は 15 桁目）');
    console.log('[PASS] Bug 32-A: Tab(n) の桁位置');
}

// --- Bug 32-A 続: Print # でも同じ ---
{
    const ev = evalVBASingle(`
Sub T()
    Dim f As Long
    f = FreeFile
    Open "tab.txt" For Output As #f
    Print #f, "X"; Tab(10); "Y"
    Close #f
End Sub`);
    ev.callProcedure('T', []);
    const content = (ev as any).fs.readFileSync('/sandbox/tab.txt').toString();
    assert.strictEqual(content.indexOf('Y') + 1, 10, 'Bug 32-A: Print # の Tab(10) でも 10 桁目');
    console.log('[PASS] Bug 32-A: Print # の Tab(n)');
}

// --- Bug 32-B: リテラルファイル番号 #1 と日付リテラルの同居 ---
{
    const ev = evalVBASingle(`
Function T() As Long
    Open "b.txt" For Output As #1
    Write #1, #2024/03/15#, 42
    Close #1
    T = 1
End Function`);
    assert.strictEqual(ev.callProcedure('T', []), 1, 'Bug 32-B: Write #1, #date# がパース・実行できる');
    console.log('[PASS] Bug 32-B: #1 + 日付リテラルの同居');
}

// --- Bug 32-C / 32-D: Write # → Input # の往復（引用符内カンマ・日付・数値） ---
{
    const ev = evalVBASingle(`
Function T() As String
    Dim f As Long, s As String, n As Long, d As Date
    f = FreeFile
    Open "w.txt" For Output As #f
    Write #f, "comma, inside", 42, #2024/03/15#
    Close #f
    f = FreeFile
    Open "w.txt" For Input As #f
    Input #f, s, n, d
    Close #f
    T = s & "|" & n & "|" & Year(d) & "/" & Month(d) & "/" & Day(d)
End Function`);
    assert.strictEqual(ev.callProcedure('T', []), 'comma, inside|42|2024/3/15',
        'Bug 32-C: 引用符内カンマ・数値・日付が Write #/Input # で往復できる');
    const content = (ev as any).fs.readFileSync('/sandbox/w.txt').toString();
    assert.strictEqual(content, '"comma, inside",42,#2024-03-15#\r\n',
        'Bug 32-D: Write # は universal date format（#yyyy-mm-dd#）+ CRLF で出力する');
    console.log('[PASS] Bug 32-C/32-D: Write #/Input # 往復と出力書式');
}

console.log('\n✅ write-input-print-zones: 全テスト通過');
