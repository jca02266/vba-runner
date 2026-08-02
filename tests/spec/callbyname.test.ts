import { evalVBASingle, assert } from '../../test-libs/test-runner';

function evalVBA(code: string): any {
    return evalVBASingle(code);
}

function runFunc(code: string, name: string, args: any[] = []): any {
    return evalVBA(code).callProcedure(name, args);
}

// --- 1. CallByName VbGet(2) / VbMethod(1) ---
{
    const code = `
    Class Employee
        Private mName As String
        Private mSalary As Long
        Public Property Get EmpName() As String
            EmpName = mName
        End Property
        Public Property Let EmpName(v As String)
            mName = v
        End Property
        Public Property Get Salary() As Long
            Salary = mSalary
        End Property
        Public Property Let Salary(v As Long)
            mSalary = v
        End Property
        Public Function Summary() As String
            Summary = mName & ":" & mSalary
        End Function
    End Class

    Function TestVbGet() As String
        Dim emp As New Employee
        emp.EmpName = "Alice"
        emp.Salary = 50000
        TestVbGet = CallByName(emp, "EmpName", 2)
    End Function

    Function TestVbMethod() As String
        Dim emp As New Employee
        emp.EmpName = "Bob"
        emp.Salary = 60000
        TestVbMethod = CallByName(emp, "Summary", 1)
    End Function
    `;
    assert.strictEqual(runFunc(code, 'TestVbGet'), 'Alice', 'CallByName VbGet(2)');
    assert.strictEqual(runFunc(code, 'TestVbMethod'), 'Bob:60000', 'CallByName VbMethod(1)');
    console.log('[PASS] CallByName VbGet/VbMethod');
}

// --- 2. Bug 27-1: CallByName VbLet(4) — Property Let ---
{
    const code = `
    Class Employee
        Private mName As String
        Private mSalary As Long
        Public Property Get EmpName() As String
            EmpName = mName
        End Property
        Public Property Let EmpName(v As String)
            mName = v
        End Property
        Public Property Get Salary() As Long
            Salary = mSalary
        End Property
        Public Property Let Salary(v As Long)
            mSalary = v
        End Property
        Public Function Summary() As String
            Summary = mName & ":" & mSalary
        End Function
    End Class

    Function TestVbLet() As String
        Dim emp As New Employee
        CallByName emp, "EmpName", 4, "Charlie"
        CallByName emp, "Salary", 4, 75000
        TestVbLet = CallByName(emp, "Summary", 1)
    End Function
    `;
    assert.strictEqual(runFunc(code, 'TestVbLet'), 'Charlie:75000', 'CallByName VbLet(4) — Bug 27-1');
    console.log('[PASS] CallByName VbLet(4) — Bug 27-1 修正確認');
}

// --- 3. CallByName VbSet(8) — Property Set ---
{
    const code = `
    Class Box
        Public Value As Long
    End Class

    Class Container
        Private mBox As Object
        Public Property Get Item() As Object
            Set Item = mBox
        End Property
        Public Property Set Item(v As Object)
            Set mBox = v
        End Property
        Public Function GetValue() As Long
            GetValue = mBox.Value
        End Function
    End Class

    Function TestVbSet() As Long
        Dim c As New Container
        Dim b As New Box
        b.Value = 99
        CallByName c, "Item", 8, b
        TestVbSet = c.GetValue()
    End Function
    `;
    assert.strictEqual(runFunc(code, 'TestVbSet'), 99, 'CallByName VbSet(8) — Property Set');
    console.log('[PASS] CallByName VbSet(8)');
}

// --- 4. CallByName VbLet — 動的プロパティ設定・読み書きラウンドトリップ ---
{
    const code = `
    Class Config
        Private mVal As String
        Public Property Get Setting() As String
            Setting = mVal
        End Property
        Public Property Let Setting(v As String)
            mVal = v
        End Property
    End Class

    Function TestRoundTrip() As String
        Dim cfg As New Config
        CallByName cfg, "Setting", 4, "dark-mode"
        TestRoundTrip = CallByName(cfg, "Setting", 2)
    End Function
    `;
    assert.strictEqual(runFunc(code, 'TestRoundTrip'), 'dark-mode', 'CallByName VbLet→VbGet ラウンドトリップ');
    console.log('[PASS] CallByName VbLet→VbGet ラウンドトリップ');
}

// Bug 167-A: CallByNameのクラス手続きByRef書き戻し
{
    const code = `
    Class Mutator
        Public Sub Apply(ByRef n As Long, ByRef label As String)
            n = n + 5
            label = label & "!"
        End Sub
    End Class

    Function TestCallByNameByRef() As String
        Dim m As New Mutator
        Dim n As Long
        Dim label As String
        n = 7
        label = "tag"
        CallByName m, "Apply", 1, n, label
        TestCallByNameByRef = CStr(n) & ":" & label
    End Function
    `;
    assert.strictEqual(runFunc(code, 'TestCallByNameByRef'), '12:tag!', 'CallByNameのByRef書き戻し');
    console.log('[PASS] Bug 167-A: CallByName class ByRef writeback');
}

// Bug 174-A: 外部COM風オブジェクトのprototypeメンバーと未定義メンバー
{
    class ExternalCom {
        private value = 'proto';
        get Value() { return this.value; }
        set Value(v: any) { this.value = String(v); }
        Add(a: any, b: any) { return Number(a) + Number(b); }
    }
    const com = new ExternalCom();
    const ev = evalVBASingle(`
Function Probe() As String
    CallByName com, "VALUE", 4, "changed"
    Probe = CStr(CallByName(com, "value", 2)) & ":" & CStr(CallByName(com, "ADD", 1, 3, 4))
End Function
`, { onPrint: () => {}, setup: e => e.set('com', com) });
    assert.strictEqual(ev.callProcedure('Probe', []), 'changed:7', '外部prototypeのGet/Let/Methodを解決する');

    const missing = evalVBASingle(`
Function Missing() As Long
    On Error Resume Next
    Dim x As Variant
    x = CallByName(com, "NoSuchProperty", 2)
    Missing = Err.Number
End Function
`, { onPrint: () => {}, setup: e => e.set('com', com) });
    assert.strictEqual(missing.callProcedure('Missing', []), 438, '未定義の外部CallByNameはError 438');
    console.log('[PASS] Bug 174-A: external CallByName prototype/missing members');
}

// Bug 373: VbLet/VbSet must not fall back to the opposite property kind
{
    const code = `
    Class OnlySet
        Private mHits As Long

        Public Property Set Target(v As Object)
            mHits = mHits + 1
        End Property

        Public Property Get Hits() As Long
            Hits = mHits
        End Property
    End Class

    Class OnlyLet
        Private mHits As Long

        Public Property Let Target(v As Variant)
            mHits = mHits + 1
        End Property

        Public Property Get Hits() As Long
            Hits = mHits
        End Property
    End Class

    Class BothKinds
        Private mLastKind As String

        Public Property Set Target(v As Object)
            mLastKind = "set"
        End Property

        Public Property Let Target(v As Variant)
            mLastKind = "let"
        End Property

        Public Property Get LastKind() As String
            LastKind = mLastKind
        End Property
    End Class

    Class Payload
    End Class

    Function VbLetDoesNotCallSet() As String
        Dim target As New OnlySet
        Dim payload As New Payload
        Dim errNo As Long
        On Error Resume Next
        CallByName target, "Target", 4, payload
        errNo = Err.Number
        On Error GoTo 0
        VbLetDoesNotCallSet = CStr(errNo) & ":" & CStr(target.Hits)
    End Function

    Function VbSetDoesNotCallLet() As String
        Dim target As New OnlyLet
        Dim payload As New Payload
        Dim errNo As Long
        On Error Resume Next
        CallByName target, "Target", 8, payload
        errNo = Err.Number
        On Error GoTo 0
        VbSetDoesNotCallLet = CStr(errNo) & ":" & CStr(target.Hits)
    End Function

    Function VbLetValuePathDoesNotCallSet() As String
        Dim target As New OnlySet
        Dim payload As New Payload
        Dim ignored As Variant
        Dim errNo As Long
        On Error Resume Next
        ignored = VBA.CallByName(target, "Target", 4, payload)
        errNo = Err.Number
        On Error GoTo 0
        VbLetValuePathDoesNotCallSet = CStr(errNo) & ":" & CStr(target.Hits)
    End Function

    Function VbSetValuePathDoesNotCallLet() As String
        Dim target As New OnlyLet
        Dim payload As New Payload
        Dim ignored As Variant
        Dim errNo As Long
        On Error Resume Next
        ignored = VBA.CallByName(target, "Target", 8, payload)
        errNo = Err.Number
        On Error GoTo 0
        VbSetValuePathDoesNotCallLet = CStr(errNo) & ":" & CStr(target.Hits)
    End Function

    Function MatchingKinds() As String
        Dim target As New BothKinds
        Dim payload As New Payload
        CallByName target, "Target", 4, 1
        MatchingKinds = target.LastKind
        CallByName target, "Target", 8, payload
        MatchingKinds = MatchingKinds & ":" & target.LastKind
    End Function
    `;
    const letMismatch = String(runFunc(code, 'VbLetDoesNotCallSet')).split(':').map(Number);
    const setMismatch = String(runFunc(code, 'VbSetDoesNotCallLet')).split(':').map(Number);
    const letValueMismatch = String(runFunc(code, 'VbLetValuePathDoesNotCallSet')).split(':').map(Number);
    const setValueMismatch = String(runFunc(code, 'VbSetValuePathDoesNotCallLet')).split(':').map(Number);
    assert.ok(letMismatch[0] !== 0, 'VbLet without Property Let must report an error');
    assert.strictEqual(letMismatch[1], 0, 'VbLet must not execute Property Set');
    assert.ok(setMismatch[0] !== 0, 'VbSet without Property Set must report an error');
    assert.strictEqual(setMismatch[1], 0, 'VbSet must not execute Property Let');
    assert.ok(letValueMismatch[0] !== 0, 'value-path VbLet without Property Let must report an error');
    assert.strictEqual(letValueMismatch[1], 0, 'value-path VbLet must not execute Property Set');
    assert.ok(setValueMismatch[0] !== 0, 'value-path VbSet without Property Set must report an error');
    assert.strictEqual(setValueMismatch[1], 0, 'value-path VbSet must not execute Property Let');
    assert.strictEqual(runFunc(code, 'MatchingKinds'), 'let:set', 'CallByName selects the exact property kind');
    console.log('[PASS] Bug 373: CallByName exact property kind');
}
