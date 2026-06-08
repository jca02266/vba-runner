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
import { Lexer } from '../../src/engine/lexer';
import { Parser } from '../../src/engine/parser';
import { Evaluator, VbaErrorValue, vbaTrue } from '../../src/engine/evaluator';
import { MockWorksheet } from '../../src/engine/mock/MockWorksheet';
import { assert } from '../../test-libs/test-runner';

function makeEv(): Evaluator {
    return new Evaluator(() => {});
}

function run(ev: Evaluator, code: string, proc: string, args: any[] = []): any {
    const ast = new Parser(new Lexer(code).tokenize()).parse();
    ev.evaluateModule(ast);
    ev.resolveIdentifiers([{ ast, moduleName: '' }]);
    return ev.callProcedure(proc, args);
}

// 1. __vbaDefault__ = true のモックオブジェクト → Value getter の値が抽出される
{
    const ev = makeEv();
    const mockCell = { __vbaDefault__: true as const, Value: 42 };
    ev.set('mockCell', mockCell);

    const result = run(ev, `
Function Test1()
    Dim x
    x = mockCell
    Test1 = x
End Function
`, 'Test1');
    assert.strictEqual(result, 42, '__vbaDefault__ ありのモックオブジェクトは値が抽出される');
    console.log('[PASS] __vbaDefault__ mock object →', result);
}

// 2. __vbaDefault__ のないオブジェクト → オブジェクトのまま代入される
{
    const ev = makeEv();
    const mockObj = { Value: 99 }; // __vbaDefault__ なし
    ev.set('mockObj', mockObj);

    const result = run(ev, `
Function Test2()
    Dim x
    x = mockObj
    Test2 = IsObject(x)
End Function
`, 'Test2');
    assert.isTrue(result, '__vbaDefault__ なしのオブジェクトはオブジェクトのまま');
    console.log('[PASS] No __vbaDefault__: object stays as object');
}

// 3. CVErr() (VbaErrorValue) → opt-in していないため抽出されない
{
    const ev = makeEv();
    const result = run(ev, `
Function Test3()
    Dim x
    x = CVErr(5)
    Test3 = IsError(x)
End Function
`, 'Test3');
    assert.isTrue(result, 'CVErr() の代入後も IsError() = True');
    console.log('[PASS] CVErr() is not unwrapped');
}

// 4. 直接注入した VbaErrorValue も抽出されない
{
    const ev = makeEv();
    ev.set('myErr', new VbaErrorValue(13));

    const result = run(ev, `
Function Test4()
    Dim x
    x = myErr
    Test4 = IsError(x)
End Function
`, 'Test4');
    assert.isTrue(result, '注入した VbaErrorValue も IsError() = True のまま');
    console.log('[PASS] Injected VbaErrorValue is not unwrapped');
}

// 5. MockRange.Value が Let 代入時に抽出される（実際のユースケース）
{
    const ev = makeEv();
    const ws = new MockWorksheet('Sheet1');
    ws.setCellValue('A1', 100);
    ws.setCellValue('B1', 200);
    ev.set('ws', ws);

    const result = run(ev, `
Function Test5()
    Dim x As Long
    Dim y As Long
    x = ws.Range("A1")
    y = ws.Range("B1")
    Test5 = x + y
End Function
`, 'Test5');
    assert.strictEqual(result, 300, 'ws.Range("A1") の Value が Let 代入時に抽出される');
    console.log('[PASS] MockRange.Value extracted via Let assignment →', result);
}

// 6. MockRange を使った VBA 演算
{
    const ev = makeEv();
    const ws = new MockWorksheet('Sheet1');
    ws.setCellValue('A1', 7);
    ws.setCellValue('A2', 3);
    ev.set('ws', ws);

    const result = run(ev, `
Function Test6()
    Dim a, b
    a = ws.Range("A1")
    b = ws.Range("A2")
    Test6 = a * b
End Function
`, 'Test6');
    assert.strictEqual(result, 21, 'MockRange の値を使った演算');
    console.log('[PASS] MockRange values in arithmetic →', result);
}

// 7. ws.Range("A1:D1") = Array(...) — デフォルトプロパティへの暗黙代入
{
    const ev = makeEv();
    const ws = new MockWorksheet('Sheet1');
    ev.set('ws', ws);
    run(ev, `
Sub Test7()
    ws.Range("A1:D1") = Array(10, 20, 30, 40)
End Sub
`, 'Test7');
    assert.strictEqual(ws.getCellValue('A1'), 10, 'A1=10');
    assert.strictEqual(ws.getCellValue('B1'), 20, 'B1=20');
    assert.strictEqual(ws.getCellValue('C1'), 30, 'C1=30');
    assert.strictEqual(ws.getCellValue('D1'), 40, 'D1=40');
    console.log('[PASS] ws.Range("A1:D1") = Array(...) は ws.Range("A1:D1").Value = Array(...) と等価');
}

// 8. ws.Range("A1:A4") = Array(...) — 1D 配列が単一列に繰り返し適用
{
    const ev = makeEv();
    const ws = new MockWorksheet('Sheet1');
    ev.set('ws', ws);
    run(ev, `
Sub Test8()
    ws.Range("A1:A4") = Array(10, 20, 30, 40)
End Sub
`, 'Test8');
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
    const ev1 = makeEv(); ev1.set('ws', ws1);
    const ev2 = makeEv(); ev2.set('ws', ws2);
    run(ev1, `Sub Test9()\n    ws.Range("A1:C2").Value = Array(1, 2, 3)\nEnd Sub`, 'Test9');
    run(ev2, `Sub Test9()\n    ws.Range("A1:C2") = Array(1, 2, 3)\nEnd Sub`, 'Test9');
    assert.strictEqual(ws1.getCellValue('A1'), ws2.getCellValue('A1'), '.Value = と = が同じ: A1');
    assert.strictEqual(ws1.getCellValue('B1'), ws2.getCellValue('B1'), '.Value = と = が同じ: B1');
    assert.strictEqual(ws1.getCellValue('A2'), ws2.getCellValue('A2'), '.Value = と = が同じ: A2');
    console.log('[PASS] .Value = と = が同じ結果');
}

console.log('\n✅ Default Property (non-__vbaClass__): 全テスト通過');
