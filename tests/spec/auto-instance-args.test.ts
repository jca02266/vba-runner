import { evalVBASingle, evalVBAModules, assert } from '../../test-libs/test-runner';

function evalVBA(code: string): any {
    return evalVBASingle(code);
}

function runFunc(code: string, name: string, args: any[] = []): any {
    return evalVBA(code).callProcedure(name, args);
}

// --- 1. VBAクラスメソッドの引数に AutoInstance を渡す (L4698) ---
// Dim x As New Cls を一度もアクセスせずにメソッドへ渡すケース。
// 内部では b.Value が MemberExpression 経由で解決されるため、
// L4698 の修正がなくても動作するが、引数を正しく解決する防衛的修正として必要。
{
    const code = `
    Class Box
        Public Value As Long
        Sub Class_Initialize()
            Value = 42
        End Sub
    End Class

    Class Receiver
        Function Unwrap(b As Object) As Long
            Unwrap = b.Value
        End Function
    End Class

    Function TestVbaMethodArg() As Long
        Dim b As New Box
        Dim r As New Receiver
        TestVbaMethodArg = r.Unwrap(b)
    End Function
    `;
    assert.strictEqual(runFunc(code, 'TestVbaMethodArg'), 42, 'VBAクラスメソッド引数: AutoInstance が解決される');
    console.log('[PASS] VBAクラスメソッドの引数に AutoInstance');
}

// --- 2. Implements ディスパッチの引数に AutoInstance を渡す (L4704) ---
{
    const code = `
    Class IWrapper
        Function Unwrap(b As Object) As Long
        End Function
    End Class

    Class Box
        Public Value As Long
        Sub Class_Initialize()
            Value = 99
        End Sub
    End Class

    Class ConcreteWrapper
        Implements IWrapper
        Function IWrapper_Unwrap(b As Object) As Long
            IWrapper_Unwrap = b.Value
        End Function
    End Class

    Function TestImplementsArg() As Long
        Dim b As New Box
        Dim w As IWrapper
        Set w = New ConcreteWrapper
        TestImplementsArg = w.Unwrap(b)
    End Function
    `;
    assert.strictEqual(runFunc(code, 'TestImplementsArg'), 99, 'Implementsディスパッチ引数: AutoInstance が解決される');
    console.log('[PASS] Implements ディスパッチの引数に AutoInstance');
}

// --- 3. CallByName の引数に AutoInstance を渡す (L4611, 修正済み) ---
// callbyname は obj.__vbaClass__ を直接参照するため、
// プレースホルダーのまま渡ると「Object doesn't support this method」になる。
// これが最初に顕在化した実害バグ。
{
    const code = `
    Class Box
        Public Value As Long
        Sub Class_Initialize()
            Value = 55
        End Sub
    End Class

    Class Receiver
        Function Unwrap(b As Object) As Long
            Unwrap = b.Value
        End Function
    End Class

    Function TestCallByNameArg() As Long
        Dim b As New Box
        Dim r As New Receiver
        TestCallByNameArg = CallByName(r, "Unwrap", VbMethod, b)
    End Function
    `;
    assert.strictEqual(runFunc(code, 'TestCallByNameArg'), 55, 'CallByName 引数: AutoInstance が解決される');
    console.log('[PASS] CallByName の引数に AutoInstance');
}

// --- 4. Dictionary の「キー」として AutoInstance を渡す (L4727 実害バグ) ---
// d.Add b, val で b がプレースホルダーのままキーに格納される。
// その後 b.Value などで b が実インスタンスに解決されると、
// 変数 b とキーに格納されたオブジェクトが別物になり
// d.Exists(b) が False を返す。
{
    const code = `
    Class Box
        Public Value As Long
        Sub Class_Initialize()
            Value = 42
        End Sub
    End Class

    Function TestDictKeyAutoInstance() As Boolean
        Dim b As New Box
        Dim d As Object
        Set d = CreateObject("Scripting.Dictionary")
        d.Add b, "found"          ' b をキーとして格納
        Dim v As Long
        v = b.Value               ' b が実インスタンスに解決される
        TestDictKeyAutoInstance = d.Exists(b)   ' 同じ実インスタンスがキーなら True
    End Function
    `;
    assert.strictEqual(runFunc(code, 'TestDictKeyAutoInstance'), -1, 'Dict キーの AutoInstance: Exists が True を返す');
    console.log('[PASS] Dictionary キーの AutoInstance が正しく解決される');
}

// --- 5. JS オブジェクトメソッドの引数（値）に AutoInstance を渡す (L4727 防衛的修正) ---
{
    const code = `
    Class Box
        Public Value As Long
        Sub Class_Initialize()
            Value = 77
        End Sub
    End Class

    Function TestDictValueArg() As Long
        Dim d As Object
        Set d = CreateObject("Scripting.Dictionary")
        Dim b As New Box
        d.Add "key", b
        Dim got As Object
        Set got = d("key")
        TestDictValueArg = got.Value
    End Function
    `;
    assert.strictEqual(runFunc(code, 'TestDictValueArg'), 77, 'Dictionary.Add 値引数: AutoInstance が解決される');
    console.log('[PASS] JS オブジェクトメソッドの値引数に AutoInstance');
}

// Bug 33-C（評価 #33）: 未使用の Dim x As New Class を ByVal で渡すと、プレースホルダーが
// callee 側で実体化され呼び出し元に反映されなかった。引数として渡す時点で
// 呼び出し元の変数に実体化するよう修正。
{
    const code = `
Public Sub PokeVal(ByVal b As Counter33C)
    b.Bump
End Sub
Public Sub PokeRef(ByRef b As Counter33C)
    b.Bump
End Sub
Public Function TestByValNew() As Long
    Dim b As New Counter33C
    PokeVal b
    PokeRef b
    TestByValNew = b.Value
End Function
`;
    const cls = `
Private m_n As Long
Public Sub Bump()
    m_n = m_n + 1
End Sub
Public Property Get Value() As Long
    Value = m_n
End Property
`;
    const ev = evalVBAModules([
        { name: 'Counter33C', code: cls, parseAsClass: 'Counter33C' },
        { name: 'Main', code },
    ]);
    assert.strictEqual(ev.callProcedure('TestByValNew', []), 2,
        'Bug 33-C: As New 変数の ByVal/ByRef 渡しで同一インスタンスが見える');
    console.log('[PASS] Bug 33-C: As New 変数の ByVal 渡し');
}

console.log('\n✅ auto-instance-args: 全テスト通過');
