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

// --- Bug 70-A: ReDim が宣言済み要素型を失い、代入時の型強制をしない ---
{
    const ev = evalVBASingle(`
        Function TestReDimTypedArray() As String
            Dim values() As Integer
            ReDim values(0 To 0)
            values(0) = 1.5
            ReDim Preserve values(0 To 1)
            values(1) = 2.5
            TestReDimTypedArray = CStr(values(0)) & ":" & CStr(values(1))
        End Function
    `);
    assert.strictEqual(ev.callProcedure('TestReDimTypedArray', []), '2:2',
        'ReDim後もInteger配列の要素代入をCIntと同様に強制変換する');
    console.log('[PASS] Bug 70-A: ReDim が型付き配列の要素型を保持する');
}

// --- Bug BN: Erase 後の動的配列に UBound を呼ぶと Error 9 ---
{
    const ev = evalVBASingle(`
        Public r1, r2, r3 As String
        Sub Test()
            Dim d() As Long
            ReDim d(2)
            d(0) = 10 : d(1) = 20 : d(2) = 30
            Erase d
            On Error Resume Next
            r1 = UBound(d)
            r3 = Err.Number & ": " & Err.Description
            On Error GoTo 0
            ReDim d(1)
            d(0) = 99
            r2 = d(0)
        End Sub
    `);
    ev.callProcedure('Test', []);
    assert.strictEqual(ev.env.get('r3'), '9: Subscript out of range', 'Erase後のUBoundはError 9');
    assert.strictEqual(ev.env.get('r2'), 99, 'Erase後にReDimすれば再利用可能');
}
console.log('[PASS] Bug BN: Erase 後の動的配列は Error 9 (UBound/LBound)');

// --- Bug 114-A: ByVal Variant 配列が呼び出し元の配列を変更する ---
{
    const ev = evalVBASingle(`
        Function TestByValVariantArray() As String
            Dim values As Variant
            values = Array("old")
            MutateCopy values
            TestByValVariantArray = values(0)
        End Function

        Private Sub MutateCopy(ByVal values As Variant)
            values(0) = "new"
        End Sub
    `);
    assert.strictEqual(ev.callProcedure('TestByValVariantArray', []), 'old',
        'ByVal Variant配列の要素変更は呼び出し元へ伝播しない');
    console.log('[PASS] Bug 114-A: ByVal Variant配列をコピーして渡す');

    const classEv = evalVBASingle(`
        Class ArrayMutator
            Public Sub Mutate(ByVal values As Variant)
                values(0) = "new"
            End Sub
        End Class
        Function TestClassByValVariantArray() As String
            Dim values As Variant
            Dim mutator As New ArrayMutator
            values = Array("old")
            mutator.Mutate values
            TestClassByValVariantArray = values(0)
        End Function
    `);
    assert.strictEqual(classEv.callProcedure('TestClassByValVariantArray', []), 'old',
        'クラスメソッドのByVal Variant配列もコピーして渡す');
    console.log('[PASS] Bug 114-A: クラスメソッドのByVal Variant配列をコピーする');

    const parenthesizedEv = evalVBASingle(`
        Function TestParenthesizedByRefArray() As Long
            Dim values(2 To 2, -1 To -1) As Long
            values(2, -1) = 7
            ChangeFirst (values)
            TestParenthesizedByRefArray = values(2, -1)
        End Function
        Private Sub ChangeFirst(ByRef values() As Long)
            values(2, -1) = 99
        End Sub
    `);
    assert.strictEqual(parenthesizedEv.callProcedure('TestParenthesizedByRefArray', []), 7,
        '括弧付きByRef配列引数は一時コピーとして渡す');
    console.log('[PASS] Bug 115-A: 括弧付きByRef配列引数をByVal化する');
}

console.log('\n✅ Array Functions: 全テスト通過');
