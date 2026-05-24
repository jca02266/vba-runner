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

console.log('\n✅ User Defined Type: 全テスト通過');
