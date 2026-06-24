/**
 * `Function Foo() As String()` のような配列戻り値型のテスト（MS-VBAL §5.3.1.6）
 *
 * 仕様バグ修正: parser.ts の戻り値型パースが `As String()` の `()`（配列マーカー）を
 * 読み捨てており、`returnType` が単なる "String" として記録されていた。これにより
 * evaluator.ts の関数戻り値代入（`GetArr = r`）処理が戻り値変数をスカラー String 型と
 * 誤認し、`coerceToDeclaredType` 経由で配列に `CStr()` を適用して配列が壊れていた。
 */
import { evalVBASingle, assert } from '../../test-libs/test-runner';

const code = `
    Public Function GetArr() As String()
        Dim r(1) As String
        r(0) = "Cola"
        r(1) = "2.00"
        GetArr = r
    End Function

    Public Function CheckIsArray() As Boolean
        CheckIsArray = IsArray(GetArr())
    End Function

    Public Function ReadFirst() As String
        Dim parsed() As String
        parsed = GetArr()
        ReadFirst = parsed(0)
    End Function
`;
const ev = evalVBASingle(code);

assert.isTrue(ev.callProcedure('CheckIsArray', []), 'GetArr() の戻り値は配列である（IsArray が True）');
assert.strictEqual(ev.callProcedure('ReadFirst', []), 'Cola', '配列戻り値を呼び出し元の配列変数で受け取り要素を読める');

console.log('[PASS] Function As String() の戻り値が配列として保持される');
console.log('\n✅ 関数の配列戻り値型: 全テスト通過');
