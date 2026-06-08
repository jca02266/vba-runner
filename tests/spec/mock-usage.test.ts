/**
 * 軽量モック実装のテストと使用例
 *
 * このファイルは、MockWorksheet を使用して
 * Excel オブジェクトに依存した VBA コードをテストする方法を示します。
 */

import { Lexer } from '../../src/engine/lexer';
import { Parser } from '../../src/engine/parser';
import { Evaluator } from '../../src/engine/evaluator';
import { MockApplication } from '../../src/engine/mock/MockExcel';
import { assert } from '../../test-libs/test-runner';

// ============================================================
// テスト 1: 単純な読み書き
// ============================================================
{
  console.log('\n[Test 1] Simple Read/Write');

  const mockApp = new MockApplication();
  const ws = mockApp.Sheets('TestSheet');

  // セットアップ
  ws.setCellValue('A1', 100);
  ws.setCellValue('B1', 200);

  // 読み取り
  const val1 = ws.Range('A1').Value;
  const val2 = ws.Range('B1').Value;

  assert.strictEqual(val1, 100, 'A1 should be 100');
  assert.strictEqual(val2, 200, 'B1 should be 200');

  console.log('[PASS] Simple Read/Write');
}

// ============================================================
// テスト 2: セルへの書き込み
// ============================================================
{
  console.log('\n[Test 2] Cell Write');

  const mockApp = new MockApplication();
  const ws = mockApp.Sheets('TestSheet');

  // 初期値を設定
  ws.Range('A1').Value = 10;

  // 変更して確認
  const original = ws.Range('A1').Value;
  assert.strictEqual(original, 10, 'Original value should be 10');

  ws.Range('A1').Value = 50;
  const updated = ws.Range('A1').Value;
  assert.strictEqual(updated, 50, 'Updated value should be 50');

  console.log('[PASS] Cell Write');
}

// ============================================================
// テスト 3: 範囲指定（配列）
// ============================================================
{
  console.log('\n[Test 3] Range (Array)');

  const mockApp = new MockApplication();
  const ws = mockApp.Sheets('TestSheet');

  // 範囲を配列で設定
  const data = [[1, 2, 3], [4, 5, 6], [7, 8, 9]];
  ws.setCellValue('A1:C3', data);

  // 範囲を取得
  const range = ws.Range('A1:C3').Value;
  assert.strictEqual(range.length, 3, 'Should have 3 rows');
  assert.strictEqual(range[0][0], 1, 'A1 should be 1');
  assert.strictEqual(range[1][1], 5, 'B2 should be 5');
  assert.strictEqual(range[2][2], 9, 'C3 should be 9');

  console.log('[PASS] Range (Array)');
}

// ============================================================
// テスト 4: VBA コード + モックの統合
// ============================================================
{
  console.log('\n[Test 4] VBA Code + Mock Integration');

  const vbaCode = `
    Function GetCellValue() As Long
        GetCellValue = Sheets("Data").Range("A1").Value
    End Function

    Sub SetCellValue()
        Sheets("Data").Range("B2").Value = 999
    End Sub

    Function SumRange() As Long
        Dim data As Variant
        data = Sheets("Data").Range("A1:A5").Value
        Dim sum As Long
        Dim i As Long
        sum = 0
        For i = 1 To UBound(data)
            Debug.Print "data(" & i & ",1)=" & data(i, 1)
            sum = sum + data(i, 1)
        Next i
        SumRange = sum
    End Function
  `;

  // モック Sheets を事前に準備
  const mockApp = new MockApplication();
  const ws = mockApp.Sheets('Data');
  ws.setCellValue('A1', 10);
  ws.setCellValue('A2', 20);
  ws.setCellValue('A3', 30);
  ws.setCellValue('A4', 40);
  ws.setCellValue('A5', 50);

  // VBA をコンパイル
  const tokens = new Lexer(vbaCode).tokenize();
  const ast = new Parser(tokens).parse();
  const ev = new Evaluator((msg: string) => {
    console.log('[DEBUG] ' + msg);
  });
  ev.evaluateModule(ast);
  ev.resolveIdentifiers([{ ast, moduleName: '' }]);

  // Sheets 関数をモックに指定（モック作成後に指定）
  ev.getGlobalEnv().set('Sheets', (name: string) => {
    const sheetObj = mockApp.Sheets(name);
    return {
      Range: (address: string) => {
        const range = sheetObj.Range(address);
        console.log('[MOCK] Range(' + address + ').Value returned:', range.Value);
        return range;
      },
    };
  });

  // テスト 1: 読み取り
  const readResult = ev.callProcedure('GetCellValue', []);
  assert.strictEqual(readResult, 10, 'GetCellValue should return 10');

  console.log('[PASS] VBA Code + Mock Integration');
}

// ============================================================
// テスト 5: 複数シートの管理
// ============================================================
{
  console.log('\n[Test 5] Multiple Sheets');

  const mockApp = new MockApplication();

  // 複数シートを作成
  mockApp.Sheets('Sheet1').setCellValue('A1', 100);
  mockApp.Sheets('Sheet2').setCellValue('A1', 200);
  mockApp.Sheets('Sheet3').setCellValue('A1', 300);

  // 各シートから読み取り
  const val1 = mockApp.Sheets('Sheet1').Range('A1').Value;
  const val2 = mockApp.Sheets('Sheet2').Range('A1').Value;
  const val3 = mockApp.Sheets('Sheet3').Range('A1').Value;

  assert.strictEqual(val1, 100, 'Sheet1 A1 should be 100');
  assert.strictEqual(val2, 200, 'Sheet2 A1 should be 200');
  assert.strictEqual(val3, 300, 'Sheet3 A1 should be 300');

  console.log('[PASS] Multiple Sheets');
}

console.log('\n✅ Mock Implementation: 全テスト通過');
