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
