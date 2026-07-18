/**
 * Bug DU: scripts/stress-on-error.ts（On Error 組み合わせストレス）で検出した 2 バグの
 * リグレッションテスト。
 *
 * 1. ネストブロック（For 等）内のエラーを GoTo ハンドラーで処理して Resume Next すると、
 *    失敗文の次ではなくブロック全体の次に飛んでいた（ループが 1 周目で中断される）。
 *    修正: executeStatements がループフレームを保ったままプロシージャ本体のハンドラーを
 *    実行し、Resume Next で同フレームの次の文へ戻る。
 *
 * 2. Err.Raise が Description 未指定時に既定メッセージを設定せず（§6.1.3.2: Number に
 *    対応する Error 関数の文字列、なければ "Application-defined or object-defined error"）、
 *    さらに素の VbaError オブジェクトのまま枠組みなしで TS 側へ escape していた。
 */
import { evalVBASingle, assert } from '../../test-libs/test-runner';

function runTrace(code: string, proc = 'Main'): string[] {
    const trace: string[] = [];
    const ev = evalVBASingle(code, { onPrint: (s: string) => trace.push(s) });
    ev.callProcedure(proc, []);
    return trace;
}

// --- 1. For ループ内エラー + GoTo ハンドラー + Resume Next → ループ継続 ---
{
    const trace = runTrace(`
Sub Main()
    On Error GoTo H
    Dim i As Long
    For i = 1 To 3
        Debug.Print "L" & i
        Err.Raise 5
        Debug.Print "E" & i
    Next i
    Debug.Print "SF"
    Exit Sub
H:
    Debug.Print "H" & Err.Number
    Resume Next
End Sub`);
    assert.strictEqual(trace.join(','), 'L1,H5,E1,L2,H5,E2,L3,H5,E3,SF',
        'Bug DU-1: ループ内エラーの Resume Next は失敗文の次に戻り全周回が継続する');
    console.log('[PASS] Bug DU-1: ループ内エラー + Resume Next でループ継続');
}

// --- 1b. Resume <label>: ネストエラーからトップレベルのラベルへ脱出 ---
{
    const trace = runTrace(`
Sub Main()
    On Error GoTo H
    Dim i As Long
    For i = 1 To 3
        Debug.Print "L" & i
        Err.Raise 5
    Next i
    Exit Sub
Fin:
    Debug.Print "FIN"
    Exit Sub
H:
    Debug.Print "H" & Err.Number
    Resume Fin
End Sub`);
    assert.strictEqual(trace.join(','), 'L1,H5,FIN',
        'Bug DU-1b: Resume <label> はループを脱出してラベルへ飛ぶ');
    console.log('[PASS] Bug DU-1b: ネストエラーから Resume <label> で脱出');
}

// --- 1c. ハンドラーが End Sub に落ちたらプロシージャ正常終了（ネストエラー起点） ---
{
    const trace = runTrace(`
Sub Main()
    On Error GoTo H
    Dim i As Long
    For i = 1 To 3
        Debug.Print "L" & i
        Err.Raise 5
    Next i
    Debug.Print "SF"
    Exit Sub
H:
    Debug.Print "H" & Err.Number
End Sub`);
    assert.strictEqual(trace.join(','), 'L1,H5',
        'Bug DU-1c: ネストエラーのハンドラーが End Sub に到達したら正常終了');
    console.log('[PASS] Bug DU-1c: ハンドラー fall-through で正常終了');
}

// --- 1d. Resume（bare）: 失敗文を再実行する（2 回目で成功するケース） ---
{
    const trace = runTrace(`
Sub Main()
    Dim fixed As Boolean
    Dim z As Variant
    On Error GoTo H
    Dim i As Long
    For i = 1 To 2
        Debug.Print "L" & i
        If Not fixed Then Err.Raise 5
        Debug.Print "E" & i
    Next i
    Exit Sub
H:
    Debug.Print "H" & Err.Number
    fixed = True
    Resume
End Sub`);
    // bare Resume が再実行するのは失敗文（単一行 If 全体）のみ。ガードの fixed が
    // True になっているので 2 回目は素通りし、L1 は再出力されない
    assert.strictEqual(trace.join(','), 'L1,H5,E1,L2,E2',
        'Bug DU-1d: bare Resume は失敗文（ループ内の単一行 If）を条件込みで再実行する');
    console.log('[PASS] Bug DU-1d: ネストエラーから bare Resume で再実行');
}

// --- 2. Err.Raise の既定 Description と枠組み ---
{
    // 既定メッセージ: VBA エラー番号に対応する文字列
    const ev = evalVBASingle(`
Function Desc5() As String
    On Error Resume Next
    Err.Raise 5
    Desc5 = Err.Description
End Function
Function DescCustom() As String
    On Error Resume Next
    Err.Raise 60000
    DescCustom = Err.Description
End Function
Sub Boom()
    Err.Raise 5
End Sub`);
    assert.strictEqual(ev.callProcedure('Desc5', []), 'Invalid procedure call or argument',
        'Bug DU-2: Err.Raise 5 の既定 Description は Error(5) の文字列');
    assert.strictEqual(ev.callProcedure('DescCustom', []), 'Application-defined or object-defined error',
        'Bug DU-2: 未知の番号の既定 Description');
    // 未処理の Err.Raise は他の実行時エラーと同じ枠組みで escape する
    assert.throwsMatch(() => ev.callProcedure('Boom', []),
        /Run-time error '5': Invalid procedure call or argument/,
        'Bug DU-2: 未処理 Err.Raise は Run-time error 形式で throw される');
    console.log('[PASS] Bug DU-2: Err.Raise の既定 Description と枠組み');
}

// --- 3. End Sub は Err を自動クリアしない（§6.1.3.2 準拠の確認） ---
{
    const ev = evalVBASingle(`
Sub Child()
    On Error Resume Next
    Err.Raise 5
End Sub
Function AfterChild() As Long
    Child
    AfterChild = Err.Number
End Function`);
    assert.strictEqual(ev.callProcedure('AfterChild', []), 5,
        '仕様確認: End Sub は Err を自動クリアしない（Exit Sub のみクリア）');
    console.log('[PASS] End Sub は Err を自動クリアしない（§6.1.3.2）');
}

console.log('\n✅ on-error-nested-resume: 全テスト通過');
