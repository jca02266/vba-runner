import { evalVBASingle, assert } from '../../test-libs/test-runner';

function run(code: string, name: string): any {
    return evalVBASingle(code).callProcedure(name, []);
}

// VarPtr: 非ゼロの Long を返す
{
    const r = run(`
        Function Test()
            Dim x As Long
            Test = VarPtr(x)
        End Function
    `, 'Test');
    assert.strictEqual(typeof r, 'number', 'VarPtr returns number');
    assert.strictEqual(r > 0, true, 'VarPtr returns positive value');
}

// StrPtr: 非ゼロの Long を返す
{
    const r = run(`
        Function Test()
            Dim s As String
            s = "hello"
            Test = StrPtr(s)
        End Function
    `, 'Test');
    assert.strictEqual(r > 0, true, 'StrPtr returns positive value');
}

// ObjPtr: 非ゼロの Long を返す
{
    const r = run(`
        Function Test()
            Dim d As Object
            Set d = CreateObject("Scripting.Dictionary")
            Test = ObjPtr(d)
        End Function
    `, 'Test');
    assert.strictEqual(r > 0, true, 'ObjPtr returns positive value');
}

// 異なる呼び出しは異なる値を返す
{
    const r = run(`
        Function Test()
            Dim a As Long, b As Long
            a = VarPtr(a)
            b = VarPtr(b)
            Test = (a <> b)
        End Function
    `, 'Test');
    assert.strictEqual(r, -1, 'VarPtr calls return distinct values');
}

console.log('✅ VarPtr/StrPtr/ObjPtr: 全テスト通過');
