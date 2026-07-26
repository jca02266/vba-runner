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

// Optional と ParamArray の併用は不可（VBA仕様）
assertCompileErrorPass1(
    `Function Sum(Optional base As Integer = 0, ParamArray vals()) As Integer\nEnd Function`,
    1, /ParamArray cannot be combined with Optional/, 'Optional と ParamArray の併用はエラー');

// ParamArray の宣言制約（末尾・Variant配列・修飾子なし）
assertCompileErrorPass1(
    `Sub F(ParamArray vals(), tail As Long)\nEnd Sub`,
    1, /ParamArray must be the last parameter/, 'ParamArray の後のパラメーターはエラー');
assertCompileErrorPass1(
    `Sub F(ByVal ParamArray vals())\nEnd Sub`,
    1, /ParamArray cannot have ByVal or ByRef/, 'ParamArray のByValはエラー');
assertCompileErrorPass1(
    `Sub F(ParamArray vals() As String)\nEnd Sub`,
    1, /ParamArray must be an array of Variant/, 'ParamArray の型指定はエラー');
assertCompileErrorPass1(
    `Sub F(ParamArray vals)\nEnd Sub`,
    1, /ParamArray must be declared as an array/, 'ParamArray の配列省略はエラー');

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
