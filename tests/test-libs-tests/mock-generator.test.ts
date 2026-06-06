import { strict as assert } from 'node:assert';
import { generateExcelMockBas, extractObjectsFromAnalyzerJson } from '../../test-libs/mock-generator';
import { evalVBAModules } from '../../test-libs/test-runner';

// generateExcelMockBas が valid な VBA を生成するかを実行で確認する
const ALL_OBJECTS = ['ActiveSheet', 'Application', 'Cells', 'Columns', 'Range', 'Rows', 'ThisWorkbook'];

{
    // 生成されたコードがパース・実行できること
    const bas = generateExcelMockBas(ALL_OBJECTS, { procName: 'TestProc', date: '2026-06-07' });

    // ヘッダーコメントが含まれること
    assert.ok(bas.includes('自動生成: vba-runner.generateMocks'), 'ヘッダーコメントがある');

    // 全オブジェクトのクラス・関数定義が含まれること
    assert.ok(bas.includes('Class MockRange'), 'MockRange クラス');
    assert.ok(bas.includes('Class MockWorksheet'), 'MockWorksheet クラス');
    assert.ok(bas.includes('Class MockRows'), 'MockRows クラス');
    assert.ok(bas.includes('Class MockColumns'), 'MockColumns クラス');
    assert.ok(bas.includes('Class MockApplication'), 'MockApplication クラス');
    assert.ok(bas.includes('Class MockWorkbook'), 'MockWorkbook クラス');
    assert.ok(bas.includes('Function ActiveSheet()'), 'ActiveSheet 関数');
    assert.ok(bas.includes('Function ThisWorkbook()'), 'ThisWorkbook 関数');

    console.log('✅ 生成内容の構造チェック OK');
}

{
    // 生成コードが VBA エンジンで実行できること
    const mockBas = generateExcelMockBas(['ActiveSheet', 'Cells', 'Rows', 'Application'], {
        date: '2026-06-07',
    });

    const testBas = `
Sub TestGenerated()
    Dim ws As Object
    Set ws = ActiveSheet()
    Dim r As Object
    Set r = ws.Cells(1, 1)
    r.Value = 42

    Dim rows As Object
    Set rows = ws.Rows()
    Debug.Print rows.Count

    Dim app As Object
    Set app = Application()
    app.ScreenUpdating = False

    Debug.Print "done"
End Sub
`;

    const ev = evalVBAModules([
        { name: 'MockExcelObjects', code: mockBas },
        { name: 'Main', code: testBas },
    ]);
    const result = ev.callProcedure('TestGenerated', []);
    assert.ok(result === undefined || result === null, 'Sub は undefined/null を返す');
    console.log('✅ 生成コードの実行 OK');
}

{
    // extractObjectsFromAnalyzerJson
    const fakeJson = {
        files: [
            {
                procedures: [
                    { excelObjectsUsed: ['ActiveSheet', 'Range'] },
                    { excelObjectsUsed: ['Application', 'Cells'] },
                ],
            },
            {
                procedures: [
                    { excelObjectsUsed: ['Rows'] },
                ],
            },
        ],
    };
    const objects = extractObjectsFromAnalyzerJson(fakeJson);
    assert.deepStrictEqual(objects, ['ActiveSheet', 'Application', 'Cells', 'Range', 'Rows']);
    console.log('✅ extractObjectsFromAnalyzerJson OK');
}

{
    // 必要なオブジェクトだけが生成されること
    const bas = generateExcelMockBas(['Application']);
    assert.ok(!bas.includes('MockRange'), 'Application のみなら MockRange は生成されない');
    assert.ok(!bas.includes('MockWorksheet'), 'Application のみなら MockWorksheet は生成されない');
    assert.ok(bas.includes('MockApplication'), 'MockApplication は含まれる');
    console.log('✅ 最小生成（Application のみ）OK');
}

console.log('All mock-generator tests passed!');
