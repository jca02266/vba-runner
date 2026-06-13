import { evalVBASingle, assert } from '../../test-libs/test-runner';

function run(code: string, name: string): any {
    return evalVBASingle(code).callProcedure(name, []);
}

// DefInt: Dim n (型なし) → Integer として動作（代入時に丸め）
{
    const code = `
        DefInt A-Z
        Function Test()
            Dim n
            n = 3.9
            Test = n
        End Function
    `;
    assert.strictEqual(run(code, 'Test'), 4, 'DefInt A-Z: Dim n → Integer (rounds 3.9 → 4)');
}

// DefStr: 型なし変数が String として動作
{
    const code = `
        DefStr S
        Function Test()
            Dim s
            s = 42
            Test = TypeName(s)
        End Function
    `;
    assert.strictEqual(run(code, 'Test'), 'String', 'DefStr S: Dim s → String');
}

// DefDbl: 数値型
{
    const code = `
        DefDbl D
        Function Test()
            Dim d
            d = 1
            Test = TypeName(d)
        End Function
    `;
    assert.strictEqual(run(code, 'Test'), 'Double', 'DefDbl D: Dim d → Double');
}

// DefStr の初期値は "" (空文字)
{
    const code = `
        DefStr S
        Function Test()
            Dim s
            Test = (s = "")
        End Function
    `;
    assert.strictEqual(run(code, 'Test'), -1, 'DefStr S: Dim s → initial value ""');
}

// DefInt の初期値は 0
{
    const code = `
        DefInt I-N
        Function Test()
            Dim n
            Test = n
        End Function
    `;
    assert.strictEqual(run(code, 'Test'), 0, 'DefInt I-N: Dim n → initial value 0');
}

// 明示的な型指定は Def-Directive より優先される
{
    const code = `
        DefInt A-Z
        Function Test()
            Dim x As Double
            x = 3.14
            Test = x
        End Function
    `;
    assert.strictEqual(run(code, 'Test'), 3.14, 'Explicit As type overrides Def-Directive');
}

// 複数レンジ: DefInt I-N, X-Z (2.9→3, 1.1→1, sum=4)
{
    const code = `
        DefInt I-N, X-Z
        Function Test()
            Dim i
            Dim x
            i = 2.9
            x = 1.1
            Test = i + x
        End Function
    `;
    assert.strictEqual(run(code, 'Test'), 4, 'DefInt multiple ranges: i=3 (rounds 2.9), x=1 (rounds 1.1) → 4');
}

// DefBool
{
    const code = `
        DefBool B
        Function Test()
            Dim b
            b = True
            Test = TypeName(b)
        End Function
    `;
    assert.strictEqual(run(code, 'Test'), 'Boolean', 'DefBool B: Dim b → Boolean');
}

// DefLng
{
    const code = `
        DefLng L
        Function Test()
            Dim l
            l = 1
            Test = TypeName(l)
        End Function
    `;
    assert.strictEqual(run(code, 'Test'), 'Long', 'DefLng L: Dim l → Long');
}

// パース確認: DefInt A-Z がエラーにならない
{
    const code = `
        DefInt A-Z
        Sub Test()
        End Sub
    `;
    let ok = false;
    try { evalVBASingle(code); ok = true; } catch {}
    assert.strictEqual(ok, true, 'DefInt A-Z parses without error');
}

console.log('✅ Def-Directive: 全テスト通過');
