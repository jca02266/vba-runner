/**
 * User Defined Type (§5.6.11) のテスト
 */
import { evalVBASingle, assert } from '../../test-libs/test-runner';

function evalVBA(code: string): any {
    return evalVBASingle(code);
}

// --- 1. Type 宣言と基本使用 ---
const typeCode = `
    Type UserInfo
        Name As String
        Age As Integer
        IsActive As Boolean
    End Type
    
    Public u As UserInfo
    
    Sub Test()
        u.Name = "Alice"
        u.Age = 30
        u.IsActive = True
    End Sub
`;
const ev1 = evalVBA(typeCode);
ev1.callProcedure('Test', []);
const u1 = ev1.env.get('u');
// JS 側ではプロパティ名は小文字で登録されているはず
assert.strictEqual(u1.name, "Alice", 'u.Name がセットされている');
assert.strictEqual(u1.age, 30, 'u.Age がセットされている');
assert.strictEqual(u1.isactive.value, -1, 'u.IsActive が True (-1) である');
console.log('[PASS] Type 宣言と基本使用');

// --- 2. ネストした Type ---
const nestedTypeCode = `
    Type Point
        X As Integer
        Y As Integer
    End Type
    Type Rect
        TopLeft As Point
        BottomRight As Point
    End Type
    
    Public r As Rect
    
    Sub Test()
        r.TopLeft.X = 10
        r.TopLeft.Y = 20
        r.BottomRight.X = 100
        r.BottomRight.Y = 200
    End Sub
`;
const ev2 = evalVBA(nestedTypeCode);
ev2.callProcedure('Test', []);
const r2 = ev2.env.get('r');
assert.strictEqual(r2.topleft.x, 10, 'r.TopLeft.X');
assert.strictEqual(r2.bottomright.y, 200, 'r.BottomRight.Y');
console.log('[PASS] ネストした Type');

// --- 3. Type 変数の初期値 ---
const typeDefaultCode = `
    Type MyType
        S As String
        I As Integer
    End Type
    Public v As MyType
    Sub Test()
    End Sub
`;
const ev3 = evalVBA(typeDefaultCode);
const v3 = ev3.env.get('v');
assert.strictEqual(v3.s, "", 'UDT メンバーの初期値 (String) は ""');
assert.strictEqual(v3.i, 0, 'UDT メンバーの初期値 (Integer) は 0');
console.log('[PASS] Type 変数の初期値');

// --- 4. 予約語をメンバー名に使用 (§5.2.3.3 reserved-name-member-dcl) ---
const reservedMemberCode = `
    Type WindowRect
        Width As Long
        Height As Long
        Left As Long
        Top As Long
        Name As String
        Value As Double
    End Type

    Public wr As WindowRect

    Sub Test()
        wr.Width = 800
        wr.Height = 600
        wr.Left = 100
        wr.Top = 50
        wr.Name = "main"
        wr.Value = 1.5
    End Sub
`;
const ev4 = evalVBA(reservedMemberCode);
ev4.callProcedure('Test', []);
const wr4 = ev4.env.get('wr');
assert.strictEqual(wr4.width,  800,    'wr.Width = 800');
assert.strictEqual(wr4.height, 600,    'wr.Height = 600');
assert.strictEqual(wr4.left,   100,    'wr.Left = 100');
assert.strictEqual(wr4.top,    50,     'wr.Top = 50');
assert.strictEqual(wr4.name,   'main', 'wr.Name = "main"');
assert.strictEqual(wr4.value,  1.5,    'wr.Value = 1.5');
console.log('[PASS] 予約語メンバー名 (Width/Height/Left/Top/Name/Value)');

// --- 5. UDT 配列: ReDim 後の要素初期化 ---
const udtArrayRedimCode = `
    Type Point
        X As Integer
        Y As Integer
    End Type

    Public pts() As Point

    Sub Test()
        ReDim pts(0 To 1)
        pts(0).X = 10
        pts(0).Y = 20
        pts(1).X = 30
        pts(1).Y = 40
    End Sub
`;
const ev5 = evalVBA(udtArrayRedimCode);
ev5.callProcedure('Test', []);
const pts5 = ev5.env.get('pts');
assert.strictEqual(typeof pts5[0], 'object', 'pts(0) は UDT インスタンス (object) である');
assert.strictEqual(typeof pts5[1], 'object', 'pts(1) は UDT インスタンス (object) である');
assert.strictEqual(pts5[0].x, 10, 'pts(0).X = 10');
assert.strictEqual(pts5[0].y, 20, 'pts(0).Y = 20');
assert.strictEqual(pts5[1].x, 30, 'pts(1).X = 30');
assert.strictEqual(pts5[1].y, 40, 'pts(1).Y = 40');
console.log('[PASS] UDT 配列: ReDim 後の要素初期化');

// --- 6. UDT 配列: 要素の独立性 (同一参照にならない) ---
const udtArrayIndepCode = `
    Type Counter
        N As Integer
    End Type

    Public cs() As Counter

    Sub Test()
        ReDim cs(0 To 1)
        cs(0).N = 99
    End Sub
`;
const ev6 = evalVBA(udtArrayIndepCode);
ev6.callProcedure('Test', []);
const cs6 = ev6.env.get('cs');
assert.strictEqual(cs6[0].n, 99, 'cs(0).N = 99');
assert.strictEqual(cs6[1].n, 0,  'cs(1).N は 0 のまま (独立したインスタンス)');
console.log('[PASS] UDT 配列: 要素の独立性');

// --- 7. UDT 配列: 初期値が UDT の既定値になっている ---
const udtArrayDefaultCode = `
    Type MyType
        S As String
        I As Integer
    End Type

    Public arr() As MyType

    Sub Test()
        ReDim arr(0 To 0)
    End Sub
`;
const ev7 = evalVBA(udtArrayDefaultCode);
ev7.callProcedure('Test', []);
const arr7 = ev7.env.get('arr');
assert.strictEqual(arr7[0].s, '',  'arr(0).S は "" (String の既定値)');
assert.strictEqual(arr7[0].i, 0,   'arr(0).I は 0 (Integer の既定値)');
console.log('[PASS] UDT 配列: 初期値が UDT の既定値');

// Test 8: B-4 — 固定サイズ UDT 配列の各要素が独立した UDT インスタンスとして初期化される
{
    const code = `
Option Explicit
Public Type Point
    X As Long
    Y As Long
End Type
Function Test8() As String
    Dim pts(2) As Point
    pts(0).X = 10
    pts(0).Y = 20
    pts(1).X = 30
    Test8 = pts(0).X & "," & pts(0).Y & "," & pts(1).X & "," & pts(1).Y
End Function
`;
    const result = evalVBA(code).callProcedure('Test8', []);
    assert.strictEqual(result, '10,20,30,0', 'Fixed-size UDT array elements are independent instances');
    console.log('[PASS] UDT 固定サイズ配列: B-4 — 各要素が独立したインスタンス');
}

// Test 9: B-4 — 固定サイズ UDT 配列の要素が共有されていない（変更が波及しない）
{
    const code = `
Option Explicit
Public Type Counter
    N As Long
End Type
Function Test9() As Long
    Dim arr(1) As Counter
    arr(0).N = 99
    Test9 = arr(1).N
End Function
`;
    const result = evalVBA(code).callProcedure('Test9', []);
    assert.strictEqual(result, 0, 'arr(1).N is unaffected by arr(0).N change (no shared reference)');
    console.log('[PASS] UDT 固定サイズ配列: B-4 — 要素間で参照が共有されていない');
}

// Test 10: 仕様バグ修正 — クラスの Private/Public フィールドが UDT 型の場合、
// 既定値が Empty のままになり Class_Initialize 内のメンバー代入が Error 91 になっていた
{
    const code = `
Type StatBlock
    Strength As Integer
End Type

Class Character
    Private mStats As StatBlock
    Public Sub Class_Initialize()
        mStats.Strength = 5
    End Sub
    Public Function GetStrength() As Integer
        GetStrength = mStats.Strength
    End Function
End Class

Function Test10() As Integer
    Dim ch As New Character
    Test10 = ch.GetStrength()
End Function
`;
    const result = evalVBA(code).callProcedure('Test10', []);
    assert.strictEqual(result, 5, 'クラスの UDT 型フィールドが instantiateType() で正しく初期化され、Class_Initialize 内のメンバー代入が機能する');
    console.log('[PASS] クラスの UDT 型フィールドが正しく初期化される（Class_Initialize 内のメンバー代入が機能）');
}

// Test 11: クラスの UDT 型フィールドは Class_Initialize がなくても既定値（0/""）で初期化される
{
    const code = `
Type Inner
    Value As Integer
End Type
Type Outer
    Name As String
    Nested As Inner
End Type

Class Widget
    Public Info As Outer
End Class

Function Test11() As String
    Dim w As New Widget
    Test11 = "[" & w.Info.Name & "][" & w.Info.Nested.Value & "]"
End Function
`;
    const result = evalVBA(code).callProcedure('Test11', []);
    assert.strictEqual(result, '[][0]', 'ネストした UDT フィールドも既定値（文字列は空、数値は0）で初期化される');
    console.log('[PASS] ネストした UDT フィールドも既定値で初期化される');
}

// Test 12: クラスの UDT 型フィールドは複数インスタンス間で参照を共有しない
{
    const code = `
Type Inner
    Value As Integer
End Type

Class Widget
    Public Info As Inner
End Class

Function Test12() As String
    Dim w1 As New Widget, w2 As New Widget
    w1.Info.Value = 100
    w2.Info.Value = 200
    Test12 = w1.Info.Value & "," & w2.Info.Value
End Function
`;
    const result = evalVBA(code).callProcedure('Test12', []);
    assert.strictEqual(result, '100,200', '複数インスタンスの UDT フィールドは独立している（参照共有なし）');
    console.log('[PASS] クラスの UDT 型フィールドは複数インスタンス間で参照を共有しない');
}

// --- Bug BT: Type メンバーに配列を持つ UDT の読み書きが Error 438 になっていた ---
{
    const code = `
Type Item
    Val As Long
    Name As String
End Type

Type Container
    Items(9) As Item
    Count As Long
End Type

Function Test13() As String
    Dim c As Container
    c.Items(0).Val = 42
    c.Items(0).Name = "First"
    c.Items(3).Val = 99
    c.Count = 2
    Test13 = c.Items(0).Val & "," & c.Items(0).Name & "," & c.Items(3).Val & "," & c.Count
End Function
`;
    const result = evalVBA(code).callProcedure('Test13', []);
    assert.strictEqual(result, '42,First,99,2', 'UDT 配列メンバーの読み書きが正常動作する');
    console.log('[PASS] Bug BT: UDT 内配列メンバー (Items(9) As Item) の読み書き');
}

// --- Bug BT: 境界を明示した Type 配列メンバー (1 To 5) ---
{
    const code = `
Type Row
    cells(1 To 5) As Long
End Type

Function Test14() As String
    Dim r As Row
    r.cells(1) = 10
    r.cells(5) = 50
    Test14 = r.cells(1) & "," & r.cells(5)
End Function
`;
    const result = evalVBA(code).callProcedure('Test14', []);
    assert.strictEqual(result, '10,50', 'Type 配列メンバーの明示的な 1 To 5 境界が正常動作する');
    console.log('[PASS] Bug BT: Type 配列メンバー 1 To N 境界の読み書き');
}

console.log('\n✅ User Defined Type: 全テスト通過');
