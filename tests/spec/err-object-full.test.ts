/**
 * Err Object (§6.1.3.2) のテスト
 */
import { evalVBASingle, assert } from '../../test-libs/test-runner';

function evalVBA(code: string): any {
    return evalVBASingle(code);
}

// --- 1. Err.Raise とプロパティ ---
const errCode = `
    Public errNum, errDesc, errSource
    Sub Test()
        On Error Resume Next
        Err.Raise 999, "MySource", "MyDescription"
        errNum = Err.Number
        errDesc = Err.Description
        errSource = Err.Source
        Err.Clear
    End Sub
`;
const ev1 = evalVBA(errCode);
ev1.callProcedure('Test', []);
assert.strictEqual(ev1.env.get('errnum'), 999, 'Err.Number');
assert.strictEqual(ev1.env.get('errdesc'), "MyDescription", 'Err.Description');
assert.strictEqual(ev1.env.get('errsource'), "MySource", 'Err.Source');
assert.strictEqual(ev1.errObj.number, 0, 'Err.Clear 後の状態');
console.log('[PASS] Err.Raise とプロパティ');

// --- 2. 暗黙的なエラー発生時の Err オブジェクト ---
const implicitErrCode = `
    Public errNum2
    Sub Test()
        On Error Resume Next
        Dim x
        x = 1 / 0
        errNum2 = Err.Number
    End Sub
`;
const ev2 = evalVBA(implicitErrCode);
ev2.callProcedure('Test', []);
assert.strictEqual(ev2.env.get('errnum2'), 11, 'Division by zero (11)');
console.log('[PASS] 暗黙的エラー時の Err オブジェクト');

// --- 3. 仕様バグ修正: 組み込み実行時エラーの Err.Description が
// "Run-time error 'N': ... (line X)" の枠組みテキストを含んでいた。
// 実 VBA の Err.Description は生のメッセージ（例: "Type mismatch"）のみを保持する。
const bareDescCode = `
    Public errDesc3
    Sub Test()
        On Error Resume Next
        Dim x As Integer
        x = "abc"
        errDesc3 = Err.Description
    End Sub
`;
const ev3 = evalVBA(bareDescCode);
ev3.callProcedure('Test', []);
assert.strictEqual(ev3.env.get('errdesc3'), 'Type mismatch', '組み込みエラーの Err.Description は枠組みなしの生のメッセージ');
console.log('[PASS] 組み込みエラーの Err.Description は枠組みテキストを含まない');

// --- 4. 仕様バグ修正: vbObjectError 定数が未実装で、
// `Err.Raise vbObjectError + n`（カスタムエラー番号のMS推奨慣用句）が
// Option Explicit 配下で未宣言変数エラーになっていた
const vbObjectErrorCode = `
    Option Explicit
    Public errNum4
    Sub Test()
        On Error Resume Next
        Err.Raise vbObjectError + 1000, "x", "custom"
        errNum4 = Err.Number
    End Sub
`;
const ev4 = evalVBA(vbObjectErrorCode);
ev4.callProcedure('Test', []);
assert.strictEqual(ev4.env.get('errnum4'), -2147221504 + 1000, 'vbObjectError + n でカスタムエラー番号を生成できる');
console.log('[PASS] vbObjectError 定数が利用できる');

console.log('\n✅ Err Object: 全テスト通過');
