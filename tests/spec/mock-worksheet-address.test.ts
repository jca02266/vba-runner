import { MockWorksheet } from '../../src/engine/mock/MockWorksheet';
import { vbaEmpty } from '../../src/engine/vba-types';
import { assert } from '../../test-libs/test-runner';

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

console.log('\n✅ mock-worksheet-address: 全テスト通過');
