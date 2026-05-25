/**
 * User Defined Type (§5.6.11) のテスト
 */
import { Lexer } from '../../src/engine/lexer';
import { Parser } from '../../src/engine/parser';
import { Evaluator } from '../../src/engine/evaluator';
import { assert } from '../../test-libs/test-runner';

function evalVBA(code: string): any {
    const tokens = new Lexer(code).tokenize();
    const ast = new Parser(tokens).parse();
    const ev = new Evaluator(console.log);
    ev.evaluate(ast);
    return ev;
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

console.log('\n✅ User Defined Type: 全テスト通過');
