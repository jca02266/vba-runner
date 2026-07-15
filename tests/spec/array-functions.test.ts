/**
 * Array Functions (UBound, LBound, Array) のテスト
 */
import { evalVBASingle, assert } from '../../test-libs/test-runner';

function evalVBA(code: string): any {
    return evalVBASingle(code);
}

// --- 1. UBound, LBound (1次元) ---
const arrayCode = `
    Public arr(1 To 5) As Integer
    Public lb, ub
    Sub Test()
        lb = LBound(arr)
        ub = UBound(arr)
    End Sub
`;
const ev1 = evalVBA(arrayCode);
ev1.callProcedure('Test', []);
assert.strictEqual(ev1.env.get('lb'), 1, 'LBound(arr) は 1');
assert.strictEqual(ev1.env.get('ub'), 5, 'UBound(arr) は 5');
console.log('[PASS] UBound, LBound (1次元)');

// --- 2. UBound, LBound (多次元) ---
const multiArrayCode = `
    Public arr(1 To 3, 0 To 10) As Integer
    Public lb1, ub1, lb2, ub2
    Sub Test()
        lb1 = LBound(arr, 1)
        ub1 = UBound(arr, 1)
        lb2 = LBound(arr, 2)
        ub2 = UBound(arr, 2)
    End Sub
`;
const ev2 = evalVBA(multiArrayCode);
ev2.callProcedure('Test', []);
assert.strictEqual(ev2.env.get('lb1'), 1, 'LBound(arr, 1)');
assert.strictEqual(ev2.env.get('ub1'), 3, 'UBound(arr, 1)');
assert.strictEqual(ev2.env.get('lb2'), 0, 'LBound(arr, 2)');
assert.strictEqual(ev2.env.get('ub2'), 10, 'UBound(arr, 2)');
console.log('[PASS] UBound, LBound (多次元)');

// --- 3. Array 関数 ---
const arrayFnCode = `
    Public v
    Sub Test()
        v = Array("A", "B", "C")
    End Sub
`;
const ev3 = evalVBA(arrayFnCode);
ev3.callProcedure('Test', []);
const v3 = ev3.env.get('v');
assert.strictEqual(Array.isArray(v3), true, 'Array() は配列を返す');
assert.strictEqual(v3.length, 3, '要素数は 3');
assert.strictEqual(v3[0], "A", 'v(0) は "A"');
console.log('[PASS] Array 関数');

// --- Bug 28-1: ReDim Preserve で UDT 配列を拡張後、新インデックス要素が Error 424 ---
{
    const udtPreserveCode = `
Type Point
    X As Long
    Y As Long
End Type

Function TestReDimPreserveUDT() As String
    Dim pts() As Point
    ReDim pts(0 To 0)
    pts(0).X = 10
    pts(0).Y = 20
    ReDim Preserve pts(0 To 2)
    pts(1).X = 30
    pts(1).Y = 40
    pts(2).X = 50
    pts(2).Y = 60
    TestReDimPreserveUDT = pts(0).X & "," & pts(1).X & "," & pts(2).X
End Function
`;
    const ev4 = evalVBA(udtPreserveCode);
    assert.strictEqual(ev4.callProcedure('TestReDimPreserveUDT', []), '10,30,50', 'ReDim Preserve UDT: 既存要素保持 + 新要素初期化');
    console.log('[PASS] Bug 28-1: ReDim Preserve UDT 配列');
}

console.log('\n✅ Array Functions: 全テスト通過');
