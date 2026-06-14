/**
 * Event Declaration (§5.2.4.3) & RaiseEvent (§5.4.2.20) のテスト
 */
import { evalVBASingle, assert, assertCompileErrorPass1 } from '../../test-libs/test-runner';

function evalVBA(code: string): any {
    return evalVBASingle(code);
}

// --- 1. Event 宣言と RaiseEvent の基本 ---
const eventCode = `
    Class MyEmitter
        Public Event StatusChanged(msg As String, code As Integer)

        Public Sub DoSomething()
            RaiseEvent StatusChanged("Processing", 1)
            RaiseEvent StatusChanged("Done", 0)
        End Sub
    End Class

    ' Client Module
    Dim WithEvents obj As MyEmitter
    Dim lastMsg
    Dim lastCode

    Sub obj_StatusChanged(msg, code)
        lastMsg = msg
        lastCode = code
    End Sub

    Sub Main()
        Set obj = New MyEmitter
        obj.DoSomething
    End Sub
`;

const ev = evalVBA(eventCode);
ev.callProcedure('Main', []);

assert.strictEqual(ev.env.get('lastmsg'), "Done", '最終的なイベントメッセージが "Done"');
assert.strictEqual(ev.env.get('lastcode'), 0, '最終的なイベントコードが 0');

console.log('[PASS] Event & RaiseEvent (WithEvents 経由)');

// --- 2. 引数なしイベント ---
const simpleEventCode = `
    Class SimpleEmitter
        Event Fired()
        Sub Fire()
            RaiseEvent Fired
        End Sub
    End Class

    Dim WithEvents s As SimpleEmitter
    Dim fireCount
    
    Sub s_Fired()
        fireCount = fireCount + 1
    End Sub
    
    Sub Test()
        fireCount = 0
        Set s = New SimpleEmitter
        s.Fire
        s.Fire
    End Sub
`;

const ev2 = evalVBA(simpleEventCode);
ev2.callProcedure('Test', []);
assert.strictEqual(ev2.env.get('firecount'), 2, 'イベントが2回発火');
console.log('[PASS] 引数なしイベント');

// --- 3. MS-VBAL §5.2.4.3: イベント名は IDENTIFIER (予約語不可) ---
// Excel の Workbook クラスは "Open" というイベントを持つが、VBA の BNF では
// event-declaration の名前に reserved-identifier は使えない。
// 実際に "Public Event Open()" はコンパイルエラーになる。
// ここでは非予約語の識別子でイベントを宣言・発火できることを確認する。
{
    const code = `
    Class WorkbookSim
        Public Event WorkbookOpen()
        Public Event BeforeClose(ByRef Cancel As Boolean)
        Public Sub OpenWorkbook()
            RaiseEvent WorkbookOpen
        End Sub
        Public Sub CloseWorkbook()
            Dim cancelFlag As Boolean
            cancelFlag = False
            RaiseEvent BeforeClose(cancelFlag)
        End Sub
    End Class

    Dim WithEvents wb As WorkbookSim
    Dim openCount As Integer
    Dim closeCount As Integer

    Private Sub wb_WorkbookOpen()
        openCount = openCount + 1
    End Sub

    Private Sub wb_BeforeClose(ByRef Cancel As Boolean)
        closeCount = closeCount + 1
        Cancel = False
    End Sub

    Sub TestEvents()
        Set wb = New WorkbookSim
        wb.OpenWorkbook
        wb.CloseWorkbook
        wb.OpenWorkbook
    End Sub
    `;
    const ev3 = evalVBA(code);
    ev3.callProcedure('TestEvents', []);
    assert.strictEqual(ev3.env.get('opencount'), 2, 'WorkbookOpen イベントが 2 回発火');
    assert.strictEqual(ev3.env.get('closecount'), 1, 'BeforeClose イベントが 1 回発火');
    console.log('[PASS] 非予約語識別子のイベント名 (WorkbookOpen/BeforeClose) が使用できる');
}

// --- 4. キーワード (Open) をイベント名にするとコンパイルエラー ---
{
    const code = `
Class Bad
    Public Event Open()
End Class
`;
    assertCompileErrorPass1(code, 3, /expected identifier after 'event'/i, 'Open をイベント名に使うとコンパイルエラー');
    console.log('[PASS] キーワード Open をイベント名にするとコンパイルエラー');
}

console.log('\n✅ Event & RaiseEvent: 全テスト通過');

