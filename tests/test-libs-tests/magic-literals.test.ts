import { Lexer } from '../../src/engine/lexer';
import { Parser } from '../../src/engine/parser';
import { findMagicLiteralsInCalls, paramName } from '../../test-libs/vba-analyzer';
import { assert } from '../../test-libs/test-runner';

function procBody(code: string): any[] {
    const tokens = new Lexer(code).tokenize();
    const ast = new Parser(tokens).parse();
    const proc = ast.body.find((s: any) => s.type === 'ProcedureDeclaration') as any;
    return proc?.body ?? [];
}

function findResults(code: string) {
    return findMagicLiteralsInCalls(procBody(code));
}

// --- paramName ---
{
    assert.strictEqual(paramName('cells', 0), 'RowIndex', 'Cells arg0');
    assert.strictEqual(paramName('cells', 1), 'ColumnIndex', 'Cells arg1');
    assert.strictEqual(paramName('Cells', 0), 'RowIndex', 'Cells 大文字');
    assert.strictEqual(paramName('range', 0), 'Cell1', 'Range arg0');
    assert.strictEqual(paramName('range', 1), 'Cell2', 'Range arg1');
    assert.strictEqual(paramName('sheets', 0), 'Index', 'Sheets arg0');
    assert.strictEqual(paramName('worksheets', 0), 'Index', 'Worksheets arg0');
    assert.strictEqual(paramName('cells.item', 0), 'RowIndex', 'Cells.Item arg0');
    assert.strictEqual(paramName('range.item', 0), 'RowIndex', 'Range.Item arg0');
    assert.strictEqual(paramName('unknown', 0), 'arg1', '未知の callee はフォールバック');
    assert.strictEqual(paramName('cells', 5), 'arg6', '範囲外はフォールバック');
    console.log('[PASS] paramName');
}

// --- 基本: 直接呼び出し ---
{
    const r = findResults(`
Sub T()
    x = Cells(3, 5).Value
End Sub`);
    assert.strictEqual(r.length, 2, 'Cells(3,5) → 2種');
    assert.strictEqual(r[0].callee, 'Cells', 'callee=Cells');
    assert.strictEqual(r[0].argIndex, 0, 'argIndex=0');
    assert.strictEqual(r[0].value, 3, 'RowIndex=3');
    assert.strictEqual(r[0].lines.length, 1, '1件');
    assert.strictEqual(r[1].value, 5, 'ColumnIndex=5');
    console.log('[PASS] 直接呼び出し Cells(3,5)');
}

// --- オブジェクト経由 ---
{
    const r = findResults(`
Sub T()
    x = ws.Cells(2, 1).Value
End Sub`);
    assert.strictEqual(r.length, 2, 'ws.Cells(2,1) → 2種');
    assert.strictEqual(r[0].callee, 'Cells', 'callee=Cells');
    assert.strictEqual(r[0].value, 2, 'RowIndex=2');
    console.log('[PASS] オブジェクト経由 ws.Cells(2,1)');
}

// --- Cells.Item 形式 ---
{
    const r = findResults(`
Sub T()
    x = Cells.Item(4, 6).Value
End Sub`);
    assert.strictEqual(r.length, 2, 'Cells.Item(4,6) → 2種');
    assert.strictEqual(r[0].callee, 'Cells.Item', 'callee=Cells.Item');
    assert.strictEqual(r[0].value, 4, 'RowIndex=4');
    assert.strictEqual(r[1].value, 6, 'ColumnIndex=6');
    console.log('[PASS] Cells.Item(4,6)');
}

// --- ws.Cells.Item 形式 ---
{
    const r = findResults(`
Sub T()
    x = ws.Cells.Item(1, 3).Value
End Sub`);
    assert.strictEqual(r[0].callee, 'Cells.Item', 'ws.Cells.Item → callee=Cells.Item');
    console.log('[PASS] ws.Cells.Item(1,3)');
}

// --- Range 直接インデックス Range("A1:B3")(2,3) ---
{
    const r = findResults(`
Sub T()
    x = Range("A1:B3")(2, 3)
End Sub`);
    const rngItems = r.filter(m => m.callee === 'Range()');
    assert.strictEqual(rngItems.length, 2, 'Range()(2,3) → 2種');
    assert.strictEqual(rngItems[0].value, 2, 'RowIndex=2');
    assert.strictEqual(rngItems[1].value, 3, 'ColumnIndex=3');
    const rngBase = r.filter(m => m.callee === 'Range');
    assert.strictEqual(rngBase.length, 1, 'Range(Cell1) も検出');
    assert.strictEqual(rngBase[0].value, 'A1:B3', 'Cell1="A1:B3"');
    console.log('[PASS] Range("A1:B3")(2,3)');
}

// --- Range().Item 形式 ---
{
    const r = findResults(`
Sub T()
    x = Range("A1:B3").Item(1, 2)
End Sub`);
    const itemResults = r.filter(m => m.callee === 'Range.Item');
    assert.strictEqual(itemResults.length, 2, 'Range.Item(1,2) → 2種');
    assert.strictEqual(itemResults[0].value, 1, 'RowIndex=1');
    console.log('[PASS] Range("A1:B3").Item(1,2)');
}

// --- Sheets / Worksheets ---
{
    const r = findResults(`
Sub T()
    Sheets("Config").Cells(1, 2).Value = 10
    Worksheets(3).Range("B2").Value = 20
End Sub`);
    const sheets = r.find(m => m.callee === 'Sheets');
    assert.strictEqual(sheets?.value, 'Config', 'Sheets(Index="Config")');
    const ws = r.find(m => m.callee === 'Worksheets');
    assert.strictEqual(ws?.value, 3, 'Worksheets(Index=3)');
    console.log('[PASS] Sheets / Worksheets');
}

// --- 同じ値の集約 ---
{
    const r = findResults(`
Sub T()
    Cells(3, 5) = "a"
    Cells(3, 5) = "b"
    Cells(3, 7) = "c"
End Sub`);
    // RowIndex=3 は3行全部に登場（L2,L3 は Cells(3,5)、L4 は Cells(3,7)）
    const row3 = r.find(m => m.callee === 'Cells' && m.argIndex === 0 && m.value === 3);
    assert.strictEqual(row3?.lines.length, 3, 'Cells(RowIndex=3) は3件に集約');
    const col5 = r.find(m => m.callee === 'Cells' && m.argIndex === 1 && m.value === 5);
    assert.strictEqual(col5?.lines.length, 2, 'Cells(ColumnIndex=5) は2件に集約');
    const col7 = r.find(m => m.callee === 'Cells' && m.argIndex === 1 && m.value === 7);
    assert.strictEqual(col7?.lines.length, 1, 'Cells(ColumnIndex=7) は1件');
    console.log('[PASS] 同じ値の集約');
}

// --- 定数・変数は検出しない ---
{
    const r = findResults(`
Sub T()
    x = Cells(ROW_HEADER, COL_STATUS).Value
    x = ws.Cells(rowNum, colNum).Value
End Sub`);
    assert.strictEqual(r.length, 0, '定数・変数は検出しない');
    console.log('[PASS] 定数・変数は検出しない');
}

// --- 左辺代入でも検出 ---
{
    const r = findResults(`
Sub T()
    Cells(2, 4) = "hello"
End Sub`);
    assert.strictEqual(r.length, 2, '左辺 Cells(2,4) も検出');
    assert.strictEqual(r[0].value, 2, 'RowIndex=2');
    console.log('[PASS] 左辺代入でも検出');
}

console.log('\n✅ magic-literals: 全テスト通過');
