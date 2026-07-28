/**
 * 非 __vbaClass__ オブジェクトの暗黙的デフォルトプロパティ (Range.Value 相当)
 *
 * VBA では `x = Range("A1")` と書いたとき Range オブジェクトの既定プロパティ
 * (Value) が暗黙的に呼ばれ、値が抽出される。
 *
 * __vbaClass__ インスタンス（VBA クラス）の場合は既存パスで処理する。
 * TypeScript から注入したモックオブジェクト（MockRange など）については
 * __vbaDefault__ = true を宣言し Value getter/setter を実装することで同等の抽出を行う。
 * evaluator は resolveObjectMemberKey(obj, 'value') で Value プロパティを解決する（読み書き対称）。
 *
 * opt-in マーカー方式の利点:
 *   - 除外リスト（VbaDate, VbaBoolean, VbaErrorValue …）が不要
 *   - 新しい内部型を追加しても壊れない
 *   - モックオブジェクト側が明示的に「既定プロパティあり」と宣言する
 *
 * このファイルは以下を検証する:
 *   1. __vbaDefault__ = true のモックオブジェクト → Let 代入時に Value getter が呼ばれる
 *   2. __vbaDefault__ のないオブジェクト → そのまま代入される
 *   3. CVErr() (VbaErrorValue) → 抽出されない (opt-in していないため)
 *   4. VbaDate / VbaBoolean → 抽出されない
 *   5. MockRange.Value が正しく抽出される（実際のユースケース）
 */
import { VbaErrorValue } from '../../src/engine/evaluator';
import { MockWorksheet } from '../../src/engine/mock/MockWorksheet';
import { VbaCurrency, VbaDecimal } from '../../src/engine/vba-types';
import { evalVBASingle, assert, vbaTrue, vbaFalse, vbaNull } from '../../test-libs/test-runner';

const VOID: EvalOptions = { onPrint: () => {} };
import type { EvalOptions } from '../../test-libs/test-runner';

function run(code: string, proc: string, setup?: (ev: any) => void, args: any[] = []): any {
    return evalVBASingle(code, { ...VOID, setup }).callProcedure(proc, args);
}

// 1. __vbaDefault__ = true のモックオブジェクト → Value getter の値が抽出される
{
    const mockCell = { __vbaDefault__: true as const, Value: 42 };
    const result = run(`
Function Test1()
    Dim x
    x = mockCell
    Test1 = x
End Function
`, 'Test1', ev => ev.set('mockCell', mockCell));
    assert.strictEqual(result, 42, '__vbaDefault__ ありのモックオブジェクトは値が抽出される');
    console.log('[PASS] __vbaDefault__ mock object →', result);
}

// 2. __vbaDefault__ のないオブジェクト → オブジェクトのまま代入される
{
    const mockObj = { Value: 99 }; // __vbaDefault__ なし
    const result = run(`
Function Test2()
    Dim x
    x = mockObj
    Test2 = IsObject(x)
End Function
`, 'Test2', ev => ev.set('mockObj', mockObj));
    assert.strictEqual(result, vbaTrue, '__vbaDefault__ なしのオブジェクトはオブジェクトのまま');
    console.log('[PASS] No __vbaDefault__: object stays as object');
}

// 3. CVErr() (VbaErrorValue) → opt-in していないため抽出されない
{
    const result = run(`
Function Test3()
    Dim x
    x = CVErr(5)
    Test3 = IsError(x)
End Function
`, 'Test3');
    assert.strictEqual(result, vbaTrue, 'CVErr() の代入後も IsError() = True');
    console.log('[PASS] CVErr() is not unwrapped');
}

// 4. 直接注入した VbaErrorValue も抽出されない
{
    const result = run(`
Function Test4()
    Dim x
    x = myErr
    Test4 = IsError(x)
End Function
`, 'Test4', ev => ev.set('myErr', new VbaErrorValue(13)));
    assert.strictEqual(result, vbaTrue, '注入した VbaErrorValue も IsError() = True のまま');
    console.log('[PASS] Injected VbaErrorValue is not unwrapped');
}

// 5. MockRange.Value が Let 代入時に抽出される（実際のユースケース）
{
    const ws = new MockWorksheet('Sheet1');
    ws.setCellValue('A1', 100);
    ws.setCellValue('B1', 200);
    const result = run(`
Function Test5()
    Dim x As Long
    Dim y As Long
    x = ws.Range("A1")
    y = ws.Range("B1")
    Test5 = x + y
End Function
`, 'Test5', ev => ev.set('ws', ws));
    assert.strictEqual(result, 300, 'ws.Range("A1") の Value が Let 代入時に抽出される');
    console.log('[PASS] MockRange.Value extracted via Let assignment →', result);
}

// Bug 155-A: 既定Valueを持つRangeをCStrへ直接渡す
{
    const ws = new MockWorksheet('Sheet1');
    ws.setCellValue('A5', 11);
    const result = run(`
Function TestCStrDefault()
    TestCStrDefault = CStr(ws.Range("A5"))
End Function
`, 'TestCStrDefault', ev => ev.set('ws', ws));
    assert.strictEqual(result, '11', '既定Valueを持つRangeはCStrで値へ変換される');
    console.log('[PASS] Bug 155-A: CStr(Range) は既定Valueを使う');
}

// Bug 156-A: 既定ValueオブジェクトをBoolean/数値型強制へ渡す
{
    const defaultNumber = { __vbaDefault__: true as const, Value: 12 };
    const result = run(`
Function TestDefaultCoercions()
    TestDefaultCoercions = CStr(CBool(v)) & ":" & CStr(CInt(v))
End Function
`, 'TestDefaultCoercions', ev => ev.set('v', defaultNumber));
    assert.strictEqual(result, 'True:12', '既定ValueオブジェクトをCBool/CIntで値へ変換する');
    console.log('[PASS] Bug 156-A: 既定ValueのBoolean/数値型強制');
}

// Bug 157-A: IsNumericも既定Valueを判定対象にする
{
    const defaultNumber = { __vbaDefault__: true as const, Value: 42 };
    const result = run(`
Function TestIsNumericDefault()
    TestIsNumericDefault = CStr(IsNumeric(v))
End Function
`, 'TestIsNumericDefault', ev => ev.set('v', defaultNumber));
    assert.strictEqual(result, 'True', 'IsNumericは既定Valueオブジェクトの値を判定する');
    console.log('[PASS] Bug 157-A: IsNumericの既定Value展開');
}

// Bug 158-A: 日付判定・CDateも既定Valueを展開する
{
    const ws = new MockWorksheet('Sheet1');
    ws.setCellValue('A6', '2024/01/15');
    const result = run(`
Function TestDateDefault()
    TestDateDefault = CStr(IsDate(ws.Range("A6"))) & ":" & CStr(Year(CDate(ws.Range("A6")))) & ":" & CStr(Year(DateValue(ws.Range("A6")))) & ":" & CStr(Year(ws.Range("A6"))) & ":" & Format(ws.Range("A6"), "yyyy")
End Function
`, 'TestDateDefault', ev => ev.set('ws', ws));
    assert.strictEqual(result, 'True:2024:2024:2024:2024', '日付関数とFormatは既定Valueの日付文字列を処理する');
    console.log('[PASS] Bug 158-A/159-A/164-A: 日付関数の既定Value展開');
}

// Bug 164-A: 日付関数群は配列・引数経路でも既定Valueを展開する
{
    const defaultDate = { __vbaDefault__: true as const, Value: '2024/01/15 14:30:45' };
    const result = run(`
Function TestDateFamily()
    TestDateFamily = CStr(Month(v)) & ":" & CStr(Day(v)) & ":" & CStr(Hour(v)) & ":" & CStr(Minute(v)) & ":" & CStr(Second(v)) & ":" & CStr(Weekday(v)) & ":" & CStr(DatePart("yyyy", v)) & ":" & CStr(Year(DateAdd("d", 1, v))) & ":" & CStr(DateDiff("d", v, DateAdd("d", 1, v))) & ":" & CStr(Hour(TimeValue(v)))
End Function
`, 'TestDateFamily', ev => ev.set('v', defaultDate));
    assert.strictEqual(result, '1:15:14:30:45:2:2024:2024:1:14', '日付関数群は既定Valueを展開する');
    console.log('[PASS] Bug 164-A: 日付関数群の既定Value展開');
}

// Bug 171-A: Currency/Decimalも数値式としてBoolean・日付へ変換する
{
    const currency = new VbaCurrency(10000n);
    const decimal = new VbaDecimal(15n, 1);
    const result = run(`
Function TestNumericObjects()
    TestNumericObjects = CStr(CBool(c)) & ":" & CStr(CBool(d)) & ":" & CStr(IsDate(c)) & ":" & CStr(Year(CDate(c))) & ":" & CStr(Year(DateValue(d)))
End Function
`, 'TestNumericObjects', ev => { ev.set('c', currency); ev.set('d', decimal); });
    assert.strictEqual(result, 'True:True:True:1899:1899', 'Currency/Decimalは数値式として型変換される');
    console.log('[PASS] Bug 171-A: Currency/Decimal boolean/date coercion');
}

// Bug 172-A: 循環する既定ValueはJSのスタックオーバーフローにならない
{
    const cyclic: any = { __vbaDefault__: true };
    cyclic.Value = cyclic;
    const result = run(`
Function TestCyclicDefault()
    On Error Resume Next
    Dim s As String
    s = CStr(v)
    TestCyclicDefault = CStr(Err.Number)
End Function
`, 'TestCyclicDefault', ev => ev.set('v', cyclic));
    assert.strictEqual(result, '13', '循環する既定ValueはError 13として処理する');
    console.log('[PASS] Bug 172-A: cyclic default Value is rejected safely');
}

// Bug 178-A: 数値書式関数も既定Valueが返すNullを直接のNullと同じ扱いにする
{
    const inner: any = { __vbaDefault__: true as const, Value: vbaNull };
    const outer: any = { __vbaDefault__: true as const, Value: inner };
    const result = run(`
Function TestNullDefault()
    TestNullDefault = FormatNumber(v) & "|" & FormatCurrency(v) & "|" & FormatPercent(v)
End Function
`, 'TestNullDefault', ev => ev.set('v', outer));
    assert.strictEqual(result, '||', '既定Valueが返すNullは数値書式関数で空文字列になる');

    const arrayResult = run(`
Function TestNullDefaultArray()
    TestNullDefaultArray = FormatNumber(values(0))
End Function
`, 'TestNullDefaultArray', ev => ev.set('values', [inner]));
    assert.strictEqual(arrayResult, '', '配列要素の既定Valueが返すNullも空文字列になる');
    console.log('[PASS] Bug 178-A: numeric format functions unwrap Null defaults');
}

// Bug 185-A: conversion functions must preserve Error 94 after default-value
// expansion returns Null, just as they do for a direct Null argument.
{
    const nullDefault: any = { __vbaDefault__: true as const, Value: vbaNull };
    const result = run(`
Function TestNullDefaultConversions() As String
    Dim e As Long, s As String
    On Error Resume Next
    x = CByte(v): e = Err.Number: Err.Clear: s = CStr(e)
    x = CInt(v): e = Err.Number: Err.Clear: s = s & ":" & CStr(e)
    x = CLng(v): e = Err.Number: Err.Clear: s = s & ":" & CStr(e)
    x = CSng(v): e = Err.Number: Err.Clear: s = s & ":" & CStr(e)
    x = CDbl(v): e = Err.Number: Err.Clear: s = s & ":" & CStr(e)
    x = CDate(v): e = Err.Number: Err.Clear: s = s & ":" & CStr(e)
    x = CDec(v): e = Err.Number: Err.Clear: s = s & ":" & CStr(e)
    x = CCur(v): e = Err.Number: Err.Clear: s = s & ":" & CStr(e)
    x = CLngLng(v): e = Err.Number: Err.Clear: s = s & ":" & CStr(e)
    x = CBool(v): e = Err.Number: Err.Clear: s = s & ":" & CStr(e)
    TestNullDefaultConversions = s
End Function
`, 'TestNullDefaultConversions', ev => ev.set('v', nullDefault));
    assert.strictEqual(result, '94:94:94:94:94:94:94:94:94:94',
        'default Value Null raises Invalid use of Null for conversions');
    console.log('[PASS] Bug 185-A: Null default conversion errors');
}

// 6. MockRange を使った VBA 演算
{
    const ws = new MockWorksheet('Sheet1');
    ws.setCellValue('A1', 7);
    ws.setCellValue('A2', 3);
    const result = run(`
Function Test6()
    Dim a, b
    a = ws.Range("A1")
    b = ws.Range("A2")
    Test6 = a * b
End Function
`, 'Test6', ev => ev.set('ws', ws));
    assert.strictEqual(result, 21, 'MockRange の値を使った演算');
    console.log('[PASS] MockRange values in arithmetic →', result);
}

// 7. ws.Range("A1:D1") = Array(...) — デフォルトプロパティへの暗黙代入
{
    const ws = new MockWorksheet('Sheet1');
    run(`
Sub Test7()
    ws.Range("A1:D1") = Array(10, 20, 30, 40)
End Sub
`, 'Test7', ev => ev.set('ws', ws));
    assert.strictEqual(ws.getCellValue('A1'), 10, 'A1=10');
    assert.strictEqual(ws.getCellValue('B1'), 20, 'B1=20');
    assert.strictEqual(ws.getCellValue('C1'), 30, 'C1=30');
    assert.strictEqual(ws.getCellValue('D1'), 40, 'D1=40');
    console.log('[PASS] ws.Range("A1:D1") = Array(...) は ws.Range("A1:D1").Value = Array(...) と等価');
}

// 8. ws.Range("A1:A4") = Array(...) — 1D 配列が単一列に繰り返し適用
{
    const ws = new MockWorksheet('Sheet1');
    run(`
Sub Test8()
    ws.Range("A1:A4") = Array(10, 20, 30, 40)
End Sub
`, 'Test8', ev => ev.set('ws', ws));
    assert.strictEqual(ws.getCellValue('A1'), 10, 'A1=10');
    assert.strictEqual(ws.getCellValue('A2'), 10, 'A2=10（繰り返し）');
    assert.strictEqual(ws.getCellValue('A3'), 10, 'A3=10（繰り返し）');
    assert.strictEqual(ws.getCellValue('A4'), 10, 'A4=10（繰り返し）');
    console.log('[PASS] ws.Range("A1:A4") = Array(10,20,30,40) → A1〜A4 すべて 10');
}

// 9. ws.Range(...).Value = と ws.Range(...) = が同じ結果
{
    const ws1 = new MockWorksheet('Sheet1');
    const ws2 = new MockWorksheet('Sheet1');
    evalVBASingle(`Sub Test9()\n    ws.Range("A1:C2").Value = Array(1, 2, 3)\nEnd Sub`,
        { ...VOID, setup: ev => ev.set('ws', ws1) }).callProcedure('Test9', []);
    evalVBASingle(`Sub Test9()\n    ws.Range("A1:C2") = Array(1, 2, 3)\nEnd Sub`,
        { ...VOID, setup: ev => ev.set('ws', ws2) }).callProcedure('Test9', []);
    assert.strictEqual(ws1.getCellValue('A1'), ws2.getCellValue('A1'), '.Value = と = が同じ: A1');
    assert.strictEqual(ws1.getCellValue('B1'), ws2.getCellValue('B1'), '.Value = と = が同じ: B1');
    assert.strictEqual(ws1.getCellValue('A2'), ws2.getCellValue('A2'), '.Value = と = が同じ: A2');
    console.log('[PASS] .Value = と = が同じ結果');
}

console.log('\n✅ Default Property (non-__vbaClass__): 全テスト通過');
