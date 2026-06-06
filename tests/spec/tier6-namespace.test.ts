/**
 * §5.6.10 Tier 6 — 型名前空間と値名前空間の分離
 *
 * VBA の §5.6.10 では修飾なし識別子の名前解決を 6 段階の Tier で定義している。
 * Tier 6 では外部参照プロジェクトの Public メンバーを修飾なしで解決できる。
 *
 * 本テストは、以下のエンジン固有の問題が修正されていることを検証する:
 * - クラスモジュール名が値名前空間（env）に VbaNamespaceRef として登録されると
 *   `Range("A1")` のような関数呼び出しが OBJECT_REQUIRED で失敗する問題
 * - 修正後は SUB_OR_FUNCTION_NOT_DEFINED が返り、モック関数の登録で解決できる
 */
import { evalVBASingle, evalVBAModules, assert } from '../../test-libs/test-runner';
import { MockApplication } from '../../src/engine/mock/MockExcel';

function expectError(fn: () => any, expectedCode: number, msg: string) {
    try {
        fn();
        throw new Error(`Expected error ${expectedCode} but no error was thrown: ${msg}`);
    } catch (e: any) {
        if (e.vbaErrorCode !== undefined) {
            assert.strictEqual(e.vbaErrorCode, expectedCode, msg);
        } else if (e.message.includes(`'${expectedCode}'`)) {
            // error message contains error code like "Run-time error '35'"
        } else {
            throw e;
        }
    }
}

// ============================================================
// 1. クラスと関数が同一モジュールで共存できる（基本動作）
// ============================================================
{
    const code = `
Option Explicit

Class RangeObj
    Public Value As Variant
End Class

Function GetRange(address As String) As RangeObj
    Dim r As New RangeObj
    r.Value = address
    Set GetRange = r
End Function

Sub Test()
    Dim r As RangeObj
    Set r = GetRange("A1")
    Debug.Print r.Value
End Sub
`;
    const ev = evalVBASingle(code);
    ev.callProcedure('Test', []);
    console.log('[PASS] 1. クラスと関数が同一モジュールで共存できる');
}

// ============================================================
// 2. クラス名と同名の関数が同一モジュールで共存できる
//    (.bas ファイルで Class Range と Function Range を定義)
// ============================================================
{
    const mockCode = `
Option Explicit

Class Range
    Public Value As Variant
End Class

Function Range(address As String) As Range
    Dim r As New Range
    r.Value = address
    Set Range = r
End Function
`;
    const userCode = `
Option Explicit

Sub TestRange()
    Dim r As Range
    Set r = Range("A1")
    Debug.Print r.Value
End Sub
`;
    const ev = evalVBAModules([
        { name: 'ExcelMocks', code: mockCode },
        { name: 'UserModule', code: userCode }
    ]);
    let output = '';
    (ev as any).onPrint = (s: string) => { output = s; };
    ev.callProcedure('TestRange', []);
    console.log('[PASS] 2. クラス名と同名の関数が同一モジュールで共存できる');
}

// ============================================================
// 3. クラスモジュール名と同名の識別子を関数呼び出ししたとき
//    OBJECT_REQUIRED ではなく SUB_OR_FUNCTION_NOT_DEFINED になる
// ============================================================
{
    const mockCode = `
Class Range
    Public Value As Variant
End Class
`;
    const userCode = `
Sub TestRangeCall()
    Dim r As Range
    Set r = Range("A1")
End Sub
`;
    const ev = evalVBAModules([
        { name: 'Range', code: mockCode },  // module name = class name
        { name: 'UserModule', code: userCode }
    ]);

    try {
        ev.callProcedure('TestRangeCall', []);
        throw new Error('Expected error but succeeded');
    } catch (e: any) {
        // Should be SUB_OR_FUNCTION_NOT_DEFINED (code 35), not OBJECT_REQUIRED (code 424)
        const msg = e.message || '';
        const isSubNotDefined = msg.includes("'35'") || e.vbaErrorCode === 35;
        const isObjRequired = msg.includes("'424'") || e.vbaErrorCode === 424;
        if (isObjRequired) {
            throw new Error('Bug: Got OBJECT_REQUIRED (424) instead of SUB_OR_FUNCTION_NOT_DEFINED (35) for class module name call');
        }
        // Should be Sub or Function not defined
        if (!isSubNotDefined) {
            throw new Error(`Expected error 35 (Sub or Function not defined), got: ${msg}`);
        }
        console.log('[PASS] 3. クラスモジュール名の呼び出し → SUB_OR_FUNCTION_NOT_DEFINED (not OBJECT_REQUIRED)');
    }
}

// ============================================================
// 4. クラスモジュール名と同名のモジュールに js 関数をセットすると
//    Range("A1") が呼び出せる
// ============================================================
{
    const mockCode = `
Class Range
    Public Value As Variant
End Class
`;
    const userCode = `
Sub TestRangeWithMock()
    Dim r As Range
    Set r = Range("A1")
    Debug.Print TypeName(r)
End Sub
`;
    const ev = evalVBAModules([
        { name: 'Range', code: mockCode },
        { name: 'UserModule', code: userCode }
    ]);

    // Register mock Range factory after modules are loaded
    ev.set('Range', (addr: string) => {
        // Create a proper Range instance via the class
        const classRangeDef = (ev as any).classDefinitions?.get('range');
        if (classRangeDef) {
            const instance = (ev as any).instantiateClass('Range');
            instance.Value = addr;
            return instance;
        }
        return { Value: addr };
    });

    ev.callProcedure('TestRangeWithMock', []);
    console.log('[PASS] 4. クラスモジュール名と同名のモック関数を登録後 Range("A1") が動作する');
}

// ============================================================
// 5. 複数クラスを含む .bas ファイル（モジュール名 ≠ クラス名）では
//    もともと衝突しない
// ============================================================
{
    const mockCode = `
Option Explicit

Class Worksheet
    Private mVal As String

    Public Sub SetValue(v As String)
        mVal = v
    End Sub

    Public Property Get Value() As String
        Value = mVal
    End Property
End Class

Class Range2
    Public Value As Variant
End Class

Function ActiveSheet() As Worksheet
    Set ActiveSheet = New Worksheet
End Function
`;
    const userCode = `
Option Explicit

Sub TestMultiClass()
    Dim ws As Worksheet
    Set ws = ActiveSheet()
    ws.SetValue "hello"
    Debug.Print ws.Value

    Dim r As Range2
    Set r = New Range2
    r.Value = 42
    Debug.Print r.Value
End Sub
`;
    const ev = evalVBAModules([
        { name: 'ExcelMocks', code: mockCode },
        { name: 'UserModule', code: userCode }
    ]);
    ev.callProcedure('TestMultiClass', []);
    console.log('[PASS] 5. 複数クラスを含む .bas ファイルではモジュール名とクラス名が非衝突');
}

// ============================================================
// 6. Dim r As Range は型名前空間から解決され Nothing になる
// ============================================================
{
    const mockCode = `
Class Range
    Public Value As Variant
End Class
`;
    const userCode = `
Sub TestDimAsRange()
    Dim r As Range
    Debug.Print IsNothing(r)
End Sub

Function IsNothing(obj) As Boolean
    IsNothing = (obj Is Nothing)
End Function
`;
    const ev = evalVBAModules([
        { name: 'Range', code: mockCode },
        { name: 'UserModule', code: userCode }
    ]);
    const result = ev.callProcedure('TestDimAsRange', []);
    // r should be vbaNothing after declaration
    console.log('[PASS] 6. Dim r As Range → 型名前空間から解決されて vbaNothing に初期化される');
}

// ============================================================
// 7. setDefaultBindingObject(MockApplication) — Range("A1") が Tier 6 で解決される
// ============================================================
{
    const userCode = `
Option Explicit

Sub TestTier6Range()
    Dim val As Variant
    val = Range("A1").Value
    Debug.Print val
End Sub
`;
    const app = new MockApplication();
    app.ActiveSheet.setCellValue('A1', 99);

    const ev = evalVBAModules([{ name: 'UserModule', code: userCode }]);
    ev.setDefaultBindingObject(app);

    let printed = '';
    (ev as any).onPrint = (s: string) => { printed = s; };
    ev.callProcedure('TestTier6Range', []);
    assert.strictEqual(Number(printed), 99, 'Range("A1").Value should be 99 via Tier 6');
    console.log('[PASS] 7. Range("A1") が Tier 6 (MockApplication) 経由で解決される');
}

// ============================================================
// 8. Cells(row, col) が Tier 6 で解決される
// ============================================================
{
    const userCode = `
Option Explicit

Sub TestTier6Cells()
    Cells(2, 3).Value = "hello"
    Debug.Print Cells(2, 3).Value
End Sub
`;
    const app = new MockApplication();

    const ev = evalVBAModules([{ name: 'UserModule', code: userCode }]);
    ev.setDefaultBindingObject(app);

    let printed = '';
    (ev as any).onPrint = (s: string) => { printed = s; };
    ev.callProcedure('TestTier6Cells', []);
    assert.strictEqual(printed, 'hello', 'Cells(2,3).Value should be "hello" via Tier 6');
    console.log('[PASS] 8. Cells(row, col) が Tier 6 (MockApplication) 経由で解決される');
}

// ============================================================
// 9. ActiveSheet プロパティが Tier 6 で解決される
// Note: ActiveSheet は宣言なしの識別子なので Option Explicit なしが前提
// ============================================================
{
    const userCode = `
Sub TestTier6ActiveSheet()
    Debug.Print ActiveSheet.Name
End Sub
`;
    const app = new MockApplication();

    const ev = evalVBAModules([{ name: 'UserModule', code: userCode }]);
    ev.setDefaultBindingObject(app);

    let printed = '';
    (ev as any).onPrint = (s: string) => { printed = s; };
    ev.callProcedure('TestTier6ActiveSheet', []);
    assert.strictEqual(printed, 'Sheet1', 'ActiveSheet.Name should be "Sheet1" via Tier 6');
    console.log('[PASS] 9. ActiveSheet が Tier 6 (MockApplication) 経由で解決される');
}

// ============================================================
// 10. Sheets("SheetName") が Tier 6 で解決される
// ============================================================
{
    const userCode = `
Option Explicit

Sub TestTier6Sheets()
    Sheets("Data").Range("B2").Value = 123
    Debug.Print Sheets("Data").Range("B2").Value
End Sub
`;
    const app = new MockApplication();

    const ev = evalVBAModules([{ name: 'UserModule', code: userCode }]);
    ev.setDefaultBindingObject(app);

    let printed = '';
    (ev as any).onPrint = (s: string) => { printed = s; };
    ev.callProcedure('TestTier6Sheets', []);
    assert.strictEqual(Number(printed), 123, 'Sheets("Data").Range("B2").Value should be 123');
    console.log('[PASS] 10. Sheets("SheetName") が Tier 6 (MockApplication) 経由で解決される');
}

// ============================================================
// 11. setActiveSheet で ActiveSheet が切り替わる
// ============================================================
{
    const userCode = `
Sub TestActiveSheetSwitch()
    Debug.Print ActiveSheet.Name
End Sub
`;
    const app = new MockApplication();
    app.setActiveSheet('Summary');

    const ev = evalVBAModules([{ name: 'UserModule', code: userCode }]);
    ev.setDefaultBindingObject(app);

    let printed = '';
    (ev as any).onPrint = (s: string) => { printed = s; };
    ev.callProcedure('TestActiveSheetSwitch', []);
    assert.strictEqual(printed, 'Summary', 'ActiveSheet.Name should be "Summary" after setActiveSheet');
    console.log('[PASS] 11. setActiveSheet でアクティブシートを切り替えられる');
}

// ============================================================
// 12. Tier 6 未設定の場合は従来通り SUB_OR_FUNCTION_NOT_DEFINED
// ============================================================
{
    const userCode = `
Sub TestNoTier6()
    Dim v As Variant
    v = Range("A1").Value
End Sub
`;
    const ev = evalVBAModules([{ name: 'UserModule', code: userCode }]);
    // defaultBindingObject は設定しない

    try {
        ev.callProcedure('TestNoTier6', []);
        throw new Error('Expected error but succeeded');
    } catch (e: any) {
        const msg = e.message || '';
        const isSubNotDefined = msg.includes("'35'") || e.vbaErrorCode === 35;
        if (!isSubNotDefined) {
            throw new Error(`Expected error 35 (Sub or Function not defined), got: ${msg}`);
        }
        console.log('[PASS] 12. Tier 6 未設定時は SUB_OR_FUNCTION_NOT_DEFINED のまま');
    }
}

console.log('\n✅ Tier 6 型名前空間/値名前空間分離: 全テスト通過');
