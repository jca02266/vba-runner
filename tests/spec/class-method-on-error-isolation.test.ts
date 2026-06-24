/**
 * クラスメソッド呼び出しが On Error の状態を呼び出し元に漏らさないことのテスト
 *
 * 仕様バグ修正: callClassMethod() が errorHandlerLabel/errorHandlingMode/
 * isInErrorHandler/lastErrorIndex を退避・復元していなかった（通常の手続き呼び出し
 * (execProcBody) は退避・復元している）。そのため、クラスメソッド内で On Error
 * ハンドラーに入った状態のまま例外が呼び出し元に伝播すると、呼び出し元の
 * executeStatements が isInErrorHandler=true を引き継いでしまい、呼び出し元自身の
 * On Error GoTo が機能せずエラーがそのまま素通りしていた。
 */
import { evalVBAModules, assert } from '../../test-libs/test-runner';

const code = `
Class Thing
    Public Function Boom() As Long
        On Error GoTo ErrHandler
        Err.Raise 600, "x", "boom"
        Exit Function
    ErrHandler:
        Err.Raise Err.Number, Err.Source, Err.Description
    End Function
End Class

Public Function CallBoom() As String
    On Error GoTo Fail
    Dim t As New Thing
    t.Boom
    CallBoom = "OK"
    Exit Function
Fail:
    CallBoom = "FAIL[" & Err.Number & "]: " & Err.Description
End Function
`;

const ev = evalVBAModules([{ name: 'Module1', code }]);
assert.strictEqual(ev.callProcedure('CallBoom', []), 'FAIL[600]: boom',
    'クラスメソッド内のハンドラーから再スローしても呼び出し元の On Error GoTo が捕捉する');

console.log('[PASS] クラスメソッド呼び出しは呼び出し元の On Error 状態を破壊しない');
console.log('\n✅ クラスメソッド On Error 状態分離: 全テスト通過');
