/**
 * クラスのオブジェクト型フィールド（クラス名・Object・Collection・WithEvents 含む）の
 * 既定値が Nothing になることのテスト（MS-VBAL: 未代入のオブジェクト参照は Nothing）
 *
 * 仕様バグ修正: createInstanceFromDef のフィールド既定値判定が文字列・数値・UDT 型しか
 * 見ておらず、オブジェクト型フィールドは Empty のままになっていた。Is Nothing 判定が
 * 常に False を返し、WithEvents バインディングの前提（未代入時は Nothing）も崩れていた。
 */
import { evalVBASingle, assert } from '../../test-libs/test-runner';

function evalVBA(code: string): any {
    return evalVBASingle(code);
}

// --- 1. クラス名型フィールドは Nothing で初期化される ---
{
    const code = `
Class Pub
End Class
Class Holder
    Public Item As Pub
End Class
Function Test1() As Boolean
    Dim h As New Holder
    Test1 = (h.Item Is Nothing)
End Function
`;
    const result = evalVBA(code).callProcedure('Test1', []);
    assert.isTrue(result, 'クラス名型フィールドの既定値は Nothing');
}
console.log('[PASS] クラス名型フィールドは Nothing で初期化される');

// --- 2. WithEvents フィールドも Nothing で初期化される ---
{
    const code = `
Class Pub
    Public Event Ping()
End Class
Class Holder
    Public WithEvents Source As Pub
End Class
Function Test2() As Boolean
    Dim h As New Holder
    Test2 = (h.Source Is Nothing)
End Function
`;
    const result = evalVBA(code).callProcedure('Test2', []);
    assert.isTrue(result, 'WithEvents フィールドの既定値も Nothing');
}
console.log('[PASS] WithEvents フィールドも Nothing で初期化される');

// --- 3. New で代入した後は Is Nothing が False になる（回帰防止） ---
{
    const code = `
Class Pub
End Class
Class Holder
    Public Item As Pub
End Class
Function Test3() As Boolean
    Dim h As New Holder
    Set h.Item = New Pub
    Test3 = (h.Item Is Nothing)
End Function
`;
    const result = evalVBA(code).callProcedure('Test3', []);
    assert.isFalse(result, 'New で代入した後は Is Nothing が False になる');
}
console.log('[PASS] New で代入した後は Is Nothing が False になる（既存挙動の維持）');

console.log('\n✅ クラスのオブジェクト型フィールドの既定値: 全テスト通過');
