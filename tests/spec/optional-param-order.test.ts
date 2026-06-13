import { evalVBASingle, assert, assertCompileErrorPass1 } from '../../test-libs/test-runner';

// §5.3.1.5: Optional パラメーターより後に非 Optional パラメーターは不可

// ❌ Optional → 非Optional はコンパイルエラー
assertCompileErrorPass1(
    `Sub F(x As Integer, Optional y As Integer, z As Integer)\nEnd Sub`,
    1, /Non-optional parameter.*z/, 'Optional の後に非Optionalはエラー');

assertCompileErrorPass1(
    `Function F(Optional x As Integer, y As Integer) As Integer\nEnd Function`,
    1, /Non-optional parameter.*y/, 'Function: Optional の後に非Optionalはエラー');

assertCompileErrorPass1(
    `Sub F(Optional a, Optional b, c)\nEnd Sub`,
    1, /Non-optional parameter.*c/, '複数Optional後の非Optionalはエラー');

// ✅ 正常なパターン（エラーにならない）

// 非Optional → Optional は OK
{
    const ev = evalVBASingle(`
        Function Test(x As Integer, Optional y As Integer = 0) As Integer
            Test = x + y
        End Function
    `);
    assert.strictEqual(ev.callProcedure('Test', [3, 2]), 5, '非Optional → Optional は正常');
    assert.strictEqual(ev.callProcedure('Test', [3]), 3, 'Optional 省略も正常');
}

// 全 Optional は OK
{
    const ev = evalVBASingle(`
        Function Test(Optional x As Integer = 1, Optional y As Integer = 2) As Integer
            Test = x + y
        End Function
    `);
    assert.strictEqual(ev.callProcedure('Test', []), 3, '全Optional は正常');
}

// Optional → ParamArray は OK (ParamArray は常に最後)
{
    const ev = evalVBASingle(`
        Function Sum(Optional base As Integer = 0, ParamArray vals())
            Dim i As Integer, total As Integer
            total = base
            For i = 0 To UBound(vals)
                total = total + vals(i)
            Next
            Sum = total
        End Function
    `);
    assert.strictEqual(ev.callProcedure('Sum', [10, 1, 2, 3]), 16, 'Optional → ParamArray は正常');
}

// ParamArray のみも OK
{
    const ev = evalVBASingle(`
        Function Sum(ParamArray vals())
            Dim i As Integer, total As Integer
            For i = 0 To UBound(vals)
                total = total + vals(i)
            Next
            Sum = total
        End Function
    `);
    assert.strictEqual(ev.callProcedure('Sum', [1, 2, 3]), 6, 'ParamArray のみは正常');
}

console.log('✅ Optional パラメーター順序チェック: 全テスト通過');
