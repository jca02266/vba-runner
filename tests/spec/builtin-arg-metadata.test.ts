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

// --- 8. Batch 2: 単純な1必須引数系の組み込み関数（テーブル駆動） ---
// 引数なし → 449 (ARGUMENT_NOT_OPTIONAL)、引数過多 → 450 (WRONG_NUMBER_OF_ARGUMENTS)
{
    const singleArgFns = [
        // 型変換
        'CByte(1)', 'CInt(1)', 'CLng(1)', 'CSng(1)', 'CDbl(1)', 'CDate(1)', 'CVDate(1)',
        'CDec(1)', 'CCur(1)', 'CLngLng(1)', 'CStr(1)', 'CBool(1)', 'CVar(1)', 'CVErr(1)',
        'Hex(1)', 'Oct(1)', 'Val("1")',
        // 情報関数
        'IsEmpty(1)', 'IsMissing(1)', 'IsNumeric(1)', 'IsDate(1)', 'IsObject(1)',
        'IsError(1)', 'IsNull(1)', 'IsArray(1)', 'VarType(1)', 'TypeName(1)',
        // 数学関数
        'Abs(1)', 'Atn(1)', 'Cos(1)', 'Exp(1)', 'Int(1)', 'Fix(1)', 'Log(1)',
        'Sgn(1)', 'Sin(1)', 'Sqr(1)', 'Tan(1)',
        // 文字列関数
        'Asc("a")', 'AscW("a")', 'Chr(65)', 'ChrW(65)', 'LCase("A")', 'Str(1)', 'UCase("a")',
        'Len("a")', 'LTrim("a")', 'RTrim("a")', 'Trim("a")', 'Space(1)', 'StrReverse("a")',
    ];
    for (const callWithArg of singleArgFns) {
        const fnName = callWithArg.slice(0, callWithArg.indexOf('('));
        expectVbaError(`Debug.Print ${fnName}()`, VbaErrorCode.ARGUMENT_NOT_OPTIONAL, `${fnName}() (引数なし)`);
        expectVbaError(`Debug.Print ${callWithArg.slice(0, -1)}, "extra")`, VbaErrorCode.WRONG_NUMBER_OF_ARGUMENTS, `${fnName}(引数過多)`);
    }
    console.log(`[PASS] Batch 2: ${singleArgFns.length} 個の単純1必須引数系関数の引数数検証`);
}

// --- 9. Batch 2: 既存の単純呼び出しは無変化（代表サンプル） ---
{
    assert.strictEqual(evalVBA(`Function F()\n F = Abs(-5)\nEnd Function`).callProcedure('F', []), 5, 'Abs(-5) = 5');
    assert.strictEqual(evalVBA(`Function F()\n F = UCase("abc")\nEnd Function`).callProcedure('F', []), 'ABC', 'UCase("abc")');
    assert.strictEqual(evalVBA(`Function F()\n F = IsNumeric("123")\nEnd Function`).callProcedure('F', []), -1, 'IsNumeric("123")');
    assert.strictEqual(evalVBA(`Function F()\n F = Round(3.14159, 2)\nEnd Function`).callProcedure('F', []), 3.14, 'Round(3.14159, 2)（必須+末尾Optional）');
    console.log('[PASS] Batch 2: 既存呼び出しの結果は無変化');
}

// --- 10. Batch 2: Round（必須1 + 末尾Optional1）の引数数エラー ---
{
    expectVbaError(`Debug.Print Round()`, VbaErrorCode.ARGUMENT_NOT_OPTIONAL, 'Round() (引数なし)');
    expectVbaError(`Debug.Print Round(1, 2, 3)`, VbaErrorCode.WRONG_NUMBER_OF_ARGUMENTS, 'Round(1,2,3) (引数過多)');
    console.log('[PASS] Round の引数数エラー (449/450)');
}

// --- 11. Batch 3: Mid/Left/Right/Format の名前付き引数（順序非依存） ---
{
    const mid1 = evalVBA(`Function F()\n F = Mid("Hello World", 2, 3)\nEnd Function`).callProcedure('F', []);
    const mid2 = evalVBA(`Function F()\n F = Mid(Length:=3, String:="Hello World", Start:=2)\nEnd Function`).callProcedure('F', []);
    assert.strictEqual(mid2, mid1, 'Mid の名前付き引数（順序を変えても同じ結果）');

    const left1 = evalVBA(`Function F()\n F = Left("Hello", 3)\nEnd Function`).callProcedure('F', []);
    const left2 = evalVBA(`Function F()\n F = Left(Length:=3, String:="Hello")\nEnd Function`).callProcedure('F', []);
    assert.strictEqual(left2, left1, 'Left の名前付き引数（順序を変えても同じ結果）');

    const fmt1 = evalVBA(`Function F()\n F = Format(3.14159, "0.00")\nEnd Function`).callProcedure('F', []);
    const fmt2 = evalVBA(`Function F()\n F = Format(Format:="0.00", Expression:=3.14159)\nEnd Function`).callProcedure('F', []);
    assert.strictEqual(fmt2, fmt1, 'Format の名前付き引数（順序を変えても同じ結果）');

    console.log('[PASS] Batch 3: Mid/Left/Format 名前付き引数（順序非依存）');
}

// --- 12. Batch 3: FV の名前付き引数（順序非依存・Optional 省略） ---
{
    const fv1 = evalVBA(`Function F()\n F = FV(0.01, 12, -100)\nEnd Function`).callProcedure('F', []);
    const fv2 = evalVBA(`Function F()\n F = FV(Pmt:=-100, Rate:=0.01, NPer:=12)\nEnd Function`).callProcedure('F', []);
    assert.strictEqual(fv2, fv1, 'FV の名前付き引数（PV/Type 省略・順序非依存）');
    console.log('[PASS] Batch 3: FV 名前付き引数');
}

// --- 13. Batch 3: 引数数エラー（代表サンプル） ---
{
    expectVbaError(`Debug.Print Left()`, VbaErrorCode.ARGUMENT_NOT_OPTIONAL, 'Left() (引数なし)');
    expectVbaError(`Debug.Print Left("a", 1, 2)`, VbaErrorCode.WRONG_NUMBER_OF_ARGUMENTS, 'Left(引数過多)');
    expectVbaError(`Debug.Print Mid("a")`, VbaErrorCode.ARGUMENT_NOT_OPTIONAL, 'Mid(引数不足)');
    expectVbaError(`Debug.Print Format()`, VbaErrorCode.ARGUMENT_NOT_OPTIONAL, 'Format() (引数なし)');
    expectVbaError(`Debug.Print MsgBox()`, VbaErrorCode.ARGUMENT_NOT_OPTIONAL, 'MsgBox() (引数なし)');
    expectVbaError(`Debug.Print FV(1, 2)`, VbaErrorCode.ARGUMENT_NOT_OPTIONAL, 'FV(引数不足)');
    console.log('[PASS] Batch 3: 引数数エラー (449/450)');
}

// --- 14. Batch 3: 既存の位置引数呼び出しは無変化（代表サンプル） ---
{
    assert.strictEqual(evalVBA(`Function F()\n F = MsgBox("hi")\nEnd Function`).callProcedure('F', []), 1, 'MsgBox("hi") は今までどおり 1 を返す');
    assert.strictEqual(evalVBA(`Function F()\n F = Right("Hello", 3)\nEnd Function`).callProcedure('F', []), 'llo', 'Right("Hello", 3) = "llo"');
    assert.strictEqual(evalVBA(`Function F()\n F = StrComp("a", "b")\nEnd Function`).callProcedure('F', []), -1, 'StrComp("a", "b") = -1');
    console.log('[PASS] Batch 3: 既存の位置引数呼び出しは無変化');
}

// --- 15. Batch 4: DateSerial/GetSetting の名前付き引数（順序非依存） ---
{
    const ds1 = evalVBA(`Function F()\n F = DateSerial(2024, 3, 15)\nEnd Function`).callProcedure('F', []);
    const ds2 = evalVBA(`Function F()\n F = DateSerial(Day:=15, Year:=2024, Month:=3)\nEnd Function`).callProcedure('F', []);
    assert.strictEqual(ds2.value, ds1.value, 'DateSerial の名前付き引数（順序を変えても同じ結果）');

    const gs1 = evalVBA(`Function F()\n F = GetSetting("App", "Sec", "Key", "default")\nEnd Function`).callProcedure('F', []);
    const gs2 = evalVBA(`Function F()\n F = GetSetting(Default:="default", Key:="Key", Section:="Sec", AppName:="App")\nEnd Function`).callProcedure('F', []);
    assert.strictEqual(gs2, gs1, 'GetSetting の名前付き引数（順序を変えても同じ結果）');
    console.log('[PASS] Batch 4: DateSerial/GetSetting 名前付き引数（順序非依存）');
}

// --- 16. Batch 4: 引数数エラー（日時・ファイル・レジストリ・misc 代表サンプル） ---
{
    expectVbaError(`Debug.Print Year()`, VbaErrorCode.ARGUMENT_NOT_OPTIONAL, 'Year() (引数なし)');
    expectVbaError(`Debug.Print Year(Date, 1)`, VbaErrorCode.WRONG_NUMBER_OF_ARGUMENTS, 'Year(引数過多)');
    expectVbaError(`Debug.Print DateSerial(2024, 3)`, VbaErrorCode.ARGUMENT_NOT_OPTIONAL, 'DateSerial(引数不足)');
    expectVbaError(`Debug.Print SaveSetting("A", "S", "K")`, VbaErrorCode.ARGUMENT_NOT_OPTIONAL, 'SaveSetting(引数不足)');
    expectVbaError(`Debug.Print GetSetting()`, VbaErrorCode.ARGUMENT_NOT_OPTIONAL, 'GetSetting() (引数なし)');
    expectVbaError(`Debug.Print IIf(True)`, VbaErrorCode.ARGUMENT_NOT_OPTIONAL, 'IIf(引数不足)');
    expectVbaError(`Debug.Print IIf(True, 1, 2, 3)`, VbaErrorCode.WRONG_NUMBER_OF_ARGUMENTS, 'IIf(引数過多)');
    expectVbaError(`Debug.Print LBound()`, VbaErrorCode.ARGUMENT_NOT_OPTIONAL, 'LBound() (引数なし)');
    console.log('[PASS] Batch 4: 引数数エラー (449/450)');
}

// --- 17. Batch 4: ParamArray 系（Choose/Switch/Array/CallByName）の既存挙動は無変化 ---
{
    assert.strictEqual(evalVBA(`Function F()\n F = Choose(2, "a", "b", "c")\nEnd Function`).callProcedure('F', []), 'b', 'Choose(2, "a","b","c") = "b"');
    assert.strictEqual(evalVBA(`Function F()\n F = Switch(False, 1, True, 2)\nEnd Function`).callProcedure('F', []), 2, 'Switch(False,1,True,2) = 2');
    const arr = evalVBA(`Function F()\n F = Array(1, 2, 3)\nEnd Function`).callProcedure('F', []);
    assert.strictEqual(arr.length, 3, 'Array(1,2,3) の要素数は3');
    const arrEmpty = evalVBA(`Function F()\n F = Array()\nEnd Function`).callProcedure('F', []);
    assert.strictEqual(arrEmpty.length, 0, 'Array() は引数なしでも呼べる（ParamArray）');
    console.log('[PASS] Batch 4: ParamArray 系の既存挙動は無変化');
}

// --- 18. Batch 4: LBound/UBound（必須+末尾Optional）は無変化 ---
{
    const r = evalVBA(`Function F()\n Dim a(1 To 5) As Integer\n F = LBound(a) & "-" & UBound(a)\nEnd Function`).callProcedure('F', []);
    assert.strictEqual(r, '1-5', 'LBound/UBound(a) は今までどおり 1-5');
    console.log('[PASS] Batch 4: LBound/UBound 既存挙動は無変化');
}

console.log('\n=== builtin-arg-metadata: 全テスト通過 ===');
