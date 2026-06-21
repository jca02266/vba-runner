import { MockWorksheet } from '../../src/engine/mock/MockWorksheet';
import { MockApplication } from '../../src/engine/mock/MockExcel';
import { vbaEmpty } from '../../src/engine/vba-types';
import { assert, evalVBASingle } from '../../test-libs/test-runner';
import { injectExcelStub } from '../../test-libs/excel-stub';
import { VbaErrorCode } from '../../src/engine/evaluator';

function ws(): MockWorksheet {
    return new MockWorksheet('Sheet1');
}

// ---- 基本動作（既存） ----

{
    const m = ws();
    m.setCellValue('A1', 42);
    assert.strictEqual(m.getCellValue('A1'), 42, '単一セル set/get');
    assert.strictEqual(m.Range('A1').Value, 42, 'Range 単一セル');
    console.log('[PASS] 基本: 単一セル');
}

{
    const m = ws();
    m.setCellValue('A1:B2', [[1, 2], [3, 4]]);
    assert.strictEqual(m.getCellValue('A1'), 1, 'A1');
    assert.strictEqual(m.getCellValue('B2'), 4, 'B2');
    console.log('[PASS] 基本: 単純範囲');
}

// ---- 絶対参照 $ ----

{
    const m = ws();
    m.setCellValue('$A$1', 99);
    assert.strictEqual(m.getCellValue('A1'), 99, '$ 付きで設定 -> 通常セルに保存');
    m.setCellValue('B2', 88);
    const v1 = m.Range('$A$1:$B$2').Value as any[][];
    const v2 = m.Range('A1:B2').Value as any[][];
    assert.strictEqual(v1[0][0], v2[0][0], '$ 付き範囲: A1');
    assert.strictEqual(v1[1][1], v2[1][1], '$ 付き範囲: B2');
    console.log('[PASS] 絶対参照 $');
}

// ---- スピル参照 # ----

{
    const m = ws();
    m.setCellValue('A1', 55);
    assert.strictEqual(m.Range('A1#').Value, 55, 'A1# -> A1 として扱う');
    console.log('[PASS] スピル参照 #');
}

// ---- トリム参照 .:. (Excel 365) ----

{
    const m = ws();
    m.setCellValue('A1:A3', [[10], [20], [30]]);
    const v = m.Range('A1.:.A3').Value as any[][];
    assert.strictEqual(v[0][0], 10, 'A1.:.A3 -> A1:A3 として扱う');
    assert.strictEqual(v[2][0], 30, 'A1.:.A3: A3');
    console.log('[PASS] トリム参照 A1.:.A3');
}

{
    const m = ws();
    m.setCellValue('B2', 55);
    assert.strictEqual(m.Range('B2.:B2').Value, 55, 'B2.:B2 -> B2 として扱う');
    assert.strictEqual(m.Range('B2:.B2').Value, 55, 'B2:.B2 -> B2 として扱う');
    console.log('[PASS] トリム参照 .: / :. バリアント');
}

// ---- implicit intersection @ (Excel 365) ----

{
    const m = ws();
    m.setCellValue('A1', 77);
    assert.strictEqual(m.Range('@A1').Value, 77, '@A1 -> A1 として扱う');
    console.log('[PASS] implicit intersection @');
}

// ---- Union カンマ ----

{
    const m = ws();
    m.setCellValue('A1', 10);
    m.setCellValue('C3', 20);
    const v = m.Range('A1:A1,C3:C3').Value as any[][];
    assert.strictEqual(v[0][0], 10, 'Union: 1エリア目');
    assert.strictEqual(v[1][0], 20, 'Union: 2エリア目');
    console.log('[PASS] Union (カンマ)');
}

{
    const m = ws();
    m.setCellValue('A1:A2', [[1], [2]]);
    m.setCellValue('C1:C2', [[3], [4]]);
    const v = m.Range('A1:A2,C1:C2').Value as any[][];
    assert.strictEqual(v.length, 4, 'Union: 行数 = 両エリアの合計');
    assert.strictEqual(v[0][0], 1, 'Union: A1');
    assert.strictEqual(v[2][0], 3, 'Union: C1');
    console.log('[PASS] Union: 複数行エリア');
}

// ---- Union setCellValue ----

{
    const m = ws();
    m.setCellValue('A1:A2,C1:C2', 99);
    assert.strictEqual(m.getCellValue('A1'), 99, 'Union set: A1');
    assert.strictEqual(m.getCellValue('A2'), 99, 'Union set: A2');
    assert.strictEqual(m.getCellValue('C1'), 99, 'Union set: C1');
    assert.strictEqual(m.getCellValue('C2'), 99, 'Union set: C2');
    assert.strictEqual(m.getCellValue('B1'), vbaEmpty, 'Union set: B1 は未設定（実Excelと同じくEmpty）');
    console.log('[PASS] Union setCellValue');
}

// ---- Intersection スペース ----

{
    const m = ws();
    m.setCellValue('A1:C3', 1);
    m.setCellValue('B2', 99);
    // A1:C3 と B2:D4 の intersection は B2:C3
    const v = m.Range('A1:C3 B2:D4').Value as any[][];
    assert.strictEqual(v.length, 2, 'Intersection: 行数 2 (B2:C3)');
    assert.strictEqual(v[0].length, 2, 'Intersection: 列数 2');
    assert.strictEqual(v[0][0], 99, 'Intersection: B2');
    console.log('[PASS] Intersection (スペース)');
}

{
    const m = ws();
    // 重なりのない2範囲 -> 空
    const v = m.Range('A1:B2 D4:E5').Value;
    assert.strictEqual(v, 0, 'Intersection: 重なりなし -> 0');
    console.log('[PASS] Intersection: 重なりなし');
}

// ---- エラーにならない確認 ----

{
    const m = ws();
    // 不正なアドレスでも例外が出ないこと
    let threw = false;
    try {
        m.Range('').Value;
        m.Range('A1# ').Value;
        m.Range('$A$1:$Z$100,B2:C3 A1:D5').Value;
    } catch {
        threw = true;
    }
    assert.strictEqual(threw, false, '複合アドレスでも例外なし');
    console.log('[PASS] 複合アドレスでも例外なし');
}

// ---- 1D 配列での書き込み（setCellValue / Range().Value = 共通） ----

{
    // 単一列範囲: 1D 配列の col[0] が全行に繰り返される
    const m = ws();
    m.setCellValue('A1:A4', [10, 20, 30, 40]);
    assert.strictEqual(m.getCellValue('A1'), 10, '1D単一列: A1=10');
    assert.strictEqual(m.getCellValue('A2'), 10, '1D単一列: A2=10（繰り返し）');
    assert.strictEqual(m.getCellValue('A3'), 10, '1D単一列: A3=10（繰り返し）');
    assert.strictEqual(m.getCellValue('A4'), 10, '1D単一列: A4=10（繰り返し）');
    console.log('[PASS] 1D 配列 単一列 (A1:A4)');
}

{
    // 複数列範囲: 各列に1D配列の対応要素が全行に繰り返される
    const m = ws();
    m.setCellValue('A1:B4', [10, 20, 30, 40]);
    assert.strictEqual(m.getCellValue('A1'), 10, '1D複数列: A1=10');
    assert.strictEqual(m.getCellValue('B1'), 20, '1D複数列: B1=20');
    assert.strictEqual(m.getCellValue('A2'), 10, '1D複数列: A2=10（繰り返し）');
    assert.strictEqual(m.getCellValue('B2'), 20, '1D複数列: B2=20（繰り返し）');
    assert.strictEqual(m.getCellValue('A4'), 10, '1D複数列: A4=10（繰り返し）');
    assert.strictEqual(m.getCellValue('B4'), 20, '1D複数列: B4=20（繰り返し）');
    console.log('[PASS] 1D 配列 複数列 (A1:B4)');
}

{
    // Range().Value = での 1D 配列書き戻し
    const m = ws();
    m.Range('A1:A4').Value = [10, 20, 30, 40];
    assert.strictEqual(m.getCellValue('A1'), 10, '1D Range書き戻し: A1=10');
    assert.strictEqual(m.getCellValue('A2'), 10, '1D Range書き戻し: A2=10（繰り返し）');
    assert.strictEqual(m.getCellValue('A4'), 10, '1D Range書き戻し: A4=10（繰り返し）');
    console.log('[PASS] 1D 配列 Range().Value = 書き戻し');
}

// ---- Range().Value = 書き戻し ----

{
    // 単一セル（既存動作の確認）
    const m = ws();
    m.Range('A1').Value = 42;
    assert.strictEqual(m.getCellValue('A1'), 42, 'Range 単一セル書き戻し');
    console.log('[PASS] Range 単一セル書き戻し');
}

{
    // 複数セル範囲 — スカラー一括書き込み
    const m = ws();
    m.Range('A1:B2').Value = 99;
    assert.strictEqual(m.getCellValue('A1'), 99, '範囲書き戻し: A1');
    assert.strictEqual(m.getCellValue('B1'), 99, '範囲書き戻し: B1');
    assert.strictEqual(m.getCellValue('A2'), 99, '範囲書き戻し: A2');
    assert.strictEqual(m.getCellValue('B2'), 99, '範囲書き戻し: B2');
    console.log('[PASS] Range 複数セル範囲 スカラー書き戻し');
}

{
    // 複数セル範囲 — 2D 配列書き込み
    const m = ws();
    m.Range('A1:B2').Value = [[1, 2], [3, 4]];
    assert.strictEqual(m.getCellValue('A1'), 1, '配列書き戻し: A1');
    assert.strictEqual(m.getCellValue('B1'), 2, '配列書き戻し: B1');
    assert.strictEqual(m.getCellValue('A2'), 3, '配列書き戻し: A2');
    assert.strictEqual(m.getCellValue('B2'), 4, '配列書き戻し: B2');
    console.log('[PASS] Range 複数セル範囲 2D 配列書き戻し');
}

{
    // 複数セル範囲 — 書き戻し後に Range().Value で再読み込み
    const m = ws();
    m.Range('A1:B2').Value = [[10, 20], [30, 40]];
    const v = m.Range('A1:B2').Value as any[][];
    assert.strictEqual(v[0][0], 10, '書き戻し後再読み込み: A1');
    assert.strictEqual(v[1][1], 40, '書き戻し後再読み込み: B2');
    console.log('[PASS] Range 書き戻し後再読み込み');
}

{
    // Union — 書き戻し
    const m = ws();
    m.Range('A1:A2,C1:C2').Value = 77;
    assert.strictEqual(m.getCellValue('A1'), 77, 'Union 書き戻し: A1');
    assert.strictEqual(m.getCellValue('A2'), 77, 'Union 書き戻し: A2');
    assert.strictEqual(m.getCellValue('C1'), 77, 'Union 書き戻し: C1');
    assert.strictEqual(m.getCellValue('C2'), 77, 'Union 書き戻し: C2');
    assert.strictEqual(m.getCellValue('B1'), vbaEmpty, 'Union 書き戻し: B1 は未設定（実Excelと同じくEmpty）');
    console.log('[PASS] Union 書き戻し');
}

// ---- Range(Cell1, Cell2) 2引数形式 ----

{
    // 文字列アドレス2つ → 角セルを結ぶ矩形
    const m = ws();
    m.setCellValue('A1:C3', [[1, 2, 3], [4, 5, 6], [7, 8, 9]]);
    const v = m.Range('A1', 'C3').Value as any[][];
    assert.strictEqual(v.length, 3, 'Range("A1","C3"): 行数3');
    assert.strictEqual(v[0].length, 3, 'Range("A1","C3"): 列数3');
    assert.strictEqual(v[0][0], 1, 'Range("A1","C3"): A1');
    assert.strictEqual(v[2][2], 9, 'Range("A1","C3"): C3');
    console.log('[PASS] Range(Cell1, Cell2): 文字列アドレス2つ');
}

{
    // Range オブジェクト2つ（Cells() の結果）→ 角セルを結ぶ矩形
    const m = ws();
    m.setCellValue('B2:D4', 1);
    const r = m.Range(m.Cells(2, 2) as any, m.Cells(4, 4) as any);
    const v = r.Value as any[][];
    assert.strictEqual(v.length, 3, 'Range(Cells,Cells): 行数3 (B2:D4)');
    assert.strictEqual(v[0].length, 3, 'Range(Cells,Cells): 列数3 (B2:D4)');
    console.log('[PASS] Range(Cell1, Cell2): Range オブジェクト2つ（Cells の結果）');
}

{
    // 角の順序が逆（右下→左上）でも矩形は正規化される
    const m = ws();
    m.setCellValue('A1:B2', [[1, 2], [3, 4]]);
    const v1 = m.Range('A1', 'B2').Value as any[][];
    const v2 = m.Range('B2', 'A1').Value as any[][];
    assert.strictEqual(v2[0][0], v1[0][0], 'Range(右下,左上) でも矩形は正規化される: A1');
    assert.strictEqual(v2[1][1], v1[1][1], 'Range(右下,左上) でも矩形は正規化される: B2');
    console.log('[PASS] Range(Cell1, Cell2): 角の順序が逆でも正規化される');
}

{
    // Range(Cell1, Cell2).Value = 書き戻し
    const m = ws();
    m.Range('A1', 'B2').Value = [[10, 20], [30, 40]];
    assert.strictEqual(m.getCellValue('A1'), 10, 'Range(Cell1,Cell2) 書き戻し: A1');
    assert.strictEqual(m.getCellValue('B2'), 40, 'Range(Cell1,Cell2) 書き戻し: B2');
    console.log('[PASS] Range(Cell1, Cell2): 書き戻し');
}

{
    // MockApplication.Range も同じく2引数対応（ActiveSheet.Range への委譲）
    const app = new MockApplication();
    app.ActiveSheet.setCellValue('A1:B2', [[1, 2], [3, 4]]);
    const v = app.Range('A1', 'B2').Value as any[][];
    assert.strictEqual(v[1][1], 4, 'MockApplication.Range(Cell1,Cell2) も同じ動作');
    console.log('[PASS] MockApplication.Range(Cell1, Cell2)');
}

// ---- Range(Cell1, Cell2) — VBA 経由（resolveCallArgs のオーバーロード機構を通す） ----

{
    const ev = evalVBASingle(`
Function F()
    Range(Cells(1, 1), Cells(3, 3)).Value = 5
    F = Cells(1, 1).Value + Cells(3, 3).Value
End Function
`, { setup: (e) => injectExcelStub(e) });
    const r = ev.callProcedure('F', []);
    assert.strictEqual(r, 10, 'VBA: Range(Cells(1,1), Cells(3,3)) で矩形書き込み');
    console.log('[PASS] VBA: Range(Cells, Cells) 2引数（Tier 6 経由）');
}

{
    // 名前付き引数（順序を変えても同じ矩形になる）
    const ev1 = evalVBASingle(`
Function F()
    Range(Cell1:="A1", Cell2:="B2").Value = 9
    F = Range("B2").Value
End Function
`, { setup: (e) => injectExcelStub(e) });
    const ev2 = evalVBASingle(`
Function F()
    Range(Cell2:="B2", Cell1:="A1").Value = 9
    F = Range("B2").Value
End Function
`, { setup: (e) => injectExcelStub(e) });
    assert.strictEqual(ev1.callProcedure('F', []), 9, 'Range の名前付き引数（正順）');
    assert.strictEqual(ev2.callProcedure('F', []), 9, 'Range の名前付き引数（逆順）');
    console.log('[PASS] VBA: Range(Cell1:=, Cell2:=) 名前付き引数（順序非依存）');
}

{
    // 引数過多（3引数）はエラー
    let caught: any = null;
    try {
        evalVBASingle(`Sub S()\n Range("A1", "B2", "C3").Value = 1\nEnd Sub`, { setup: (e) => injectExcelStub(e) }).callProcedure('S', []);
    } catch (e: any) {
        caught = e;
    }
    assert.strictEqual(caught?.number, VbaErrorCode.WRONG_NUMBER_OF_ARGUMENTS, 'Range(3引数) は 450 エラー');
    console.log('[PASS] VBA: Range(引数過多) は引数数エラー (450)');
}

console.log('\n✅ mock-worksheet-address: 全テスト通過');
