/**
 * 組み込み関数の引数メタデータ（Batch 0/1）のテスト
 *
 * 対象: Now/Date/Time/Timer/Rnd/FreeFile/DoEvents/Error/GetObject — registerBuiltin で移行。
 *       InStr/InStrB — registerOverloadedBuiltin（VBA にはないオーバーロード機構）で移行。
 *
 * 検証する挙動:
 *   1. 必須引数を超える数の引数を渡すとコンパイルエラー（450: WRONG_NUMBER_OF_ARGUMENTS）
 *   2. 括弧無し参照（auto-call）が今までどおり動く
 *   3. InStr: 名前付き引数（:=）が VBA パラメーター名どおりに解決される（順序に依存しない）
 *   4. InStr: 引数が足りない呼び出しは 449（ARGUMENT_NOT_OPTIONAL）
 *   5. InStr: 既存の位置引数呼び出し（2/3/4引数）は変更なし
 */
import { evalVBASingle, assert } from '../../test-libs/test-runner';
import { VbaErrorCode } from '../../src/engine/evaluator';

function evalVBA(code: string): any {
    return evalVBASingle(code);
}

function expectVbaError(code: string, expectedNumber: number, label: string): void {
    let caught: any = null;
    try {
        evalVBA(code);
    } catch (e: any) {
        caught = e;
    }
    assert.strictEqual(caught !== null, true, `${label}: エラーが発生する`);
    assert.strictEqual(caught?.number, expectedNumber, `${label}: エラー番号 ${expectedNumber}（got: ${caught?.number}）`);
}

// --- 1. Now/Date/Time/Timer/DoEvents: 引数を渡すとエラー（今までは黙って無視されていた） ---
{
    expectVbaError(`Debug.Print Now(1)`, VbaErrorCode.WRONG_NUMBER_OF_ARGUMENTS, 'Now(1)');
    expectVbaError(`Debug.Print Date(1)`, VbaErrorCode.WRONG_NUMBER_OF_ARGUMENTS, 'Date(1)');
    expectVbaError(`Debug.Print Time(1)`, VbaErrorCode.WRONG_NUMBER_OF_ARGUMENTS, 'Time(1)');
    expectVbaError(`Debug.Print Timer(1)`, VbaErrorCode.WRONG_NUMBER_OF_ARGUMENTS, 'Timer(1)');
    expectVbaError(`Debug.Print DoEvents(1)`, VbaErrorCode.WRONG_NUMBER_OF_ARGUMENTS, 'DoEvents(1)');
    console.log('[PASS] Now/Date/Time/Timer/DoEvents(extra arg) → 450');
}

// --- 2. 括弧無し参照（auto-call）が今までどおり動く ---
{
    const r = evalVBA(`
Function TestAutoCall()
    TestAutoCall = (Now >= 0) And (Date >= 0) And (Time >= 0) And (Timer >= 0)
End Function
`).callProcedure('TestAutoCall', []);
    assert.strictEqual(r, -1 /* vbaTrue */, 'Now/Date/Time/Timer の括弧無し参照は今までどおり動く');
    console.log('[PASS] 括弧無し参照 (auto-call) は維持される');
}

// --- 3. Rnd/FreeFile/GetObject: Optional 引数つきで呼べる（位置引数、今までと同じ） ---
{
    const r = evalVBA(`Function F()\n F = Rnd(1) >= 0 And Rnd(1) <= 1\nEnd Function`).callProcedure('F', []);
    assert.strictEqual(r, -1, 'Rnd(1) は今までどおり 0 以上 1 以下');
    console.log('[PASS] Rnd(Number) 位置引数は維持される');
}

// --- 4. InStr: 名前付き引数が VBA パラメーター名で解決される（順序に依存しない） ---
{
    const positional = evalVBA(`Function F()\n F = InStr(1, "abcabc", "b")\nEnd Function`).callProcedure('F', []);
    const namedInOrder = evalVBA(`Function F()\n F = InStr(Start:=1, String1:="abcabc", String2:="b")\nEnd Function`).callProcedure('F', []);
    const namedReordered = evalVBA(`Function F()\n F = InStr(String2:="b", String1:="abcabc", Start:=1)\nEnd Function`).callProcedure('F', []);
    assert.strictEqual(namedInOrder, positional, 'InStr(Start:=1, String1:=..., String2:=...) は位置引数版と同じ結果');
    assert.strictEqual(namedReordered, positional, '名前付き引数の並び順を変えても結果は変わらない');
    console.log('[PASS] InStr 名前付き引数（順序非依存）');
}

// --- 5. InStr: 2引数の名前付き呼び出し（Start/Compare を省略） ---
{
    const positional = evalVBA(`Function F()\n F = InStr("Hello World", "World")\nEnd Function`).callProcedure('F', []);
    const named = evalVBA(`Function F()\n F = InStr(String2:="World", String1:="Hello World")\nEnd Function`).callProcedure('F', []);
    assert.strictEqual(named, positional, 'InStr(String1/String2 のみ名前付き) は2引数版と同じ結果');
    console.log('[PASS] InStr 2引数オーバーロードの名前付き呼び出し');
}

// --- 6. InStr: 引数が足りない／多すぎる呼び出しはエラー ---
{
    expectVbaError(`Debug.Print InStr()`, VbaErrorCode.ARGUMENT_NOT_OPTIONAL, 'InStr()');
    expectVbaError(`Debug.Print InStr("a")`, VbaErrorCode.ARGUMENT_NOT_OPTIONAL, 'InStr("a")');
    expectVbaError(`Debug.Print InStr(1, "a", "b", 1, "extra")`, VbaErrorCode.WRONG_NUMBER_OF_ARGUMENTS, 'InStr(5引数)');
    console.log('[PASS] InStr 引数個数エラー (449/450)');
}

// --- 7. InStr: 既存の位置引数呼び出し（2/3/4引数）は変更なし ---
{
    assert.strictEqual(evalVBA(`Function F()\n F = InStr("Hello World", "World")\nEnd Function`).callProcedure('F', []), 7, '2引数版');
    assert.strictEqual(evalVBA(`Function F()\n F = InStr(1, "abcabc", "b")\nEnd Function`).callProcedure('F', []), 2, '3引数版');
    assert.strictEqual(evalVBA(`Function F()\n F = InStr(1, "abcabc", "b", 0)\nEnd Function`).callProcedure('F', []), 2, '4引数版');
    console.log('[PASS] InStr 既存の位置引数呼び出しは無変化');
}

console.log('\n=== builtin-arg-metadata: 全テスト通過 ===');
