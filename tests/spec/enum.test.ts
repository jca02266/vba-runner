/**
 * Enum 宣言のテスト (§5.2.3.4)
 *
 * 検証する挙動:
 *   - 明示的な値の指定
 *   - 値を省略した場合の自動インクリメント
 *   - メンバ名による直接アクセス（Red 等）
 *   - Enum 名経由のアクセス（Color.Red 等）
 */
import { evalVBASingle, assert } from '../../test-libs/test-runner';

function evalVBA(code: string): any {
    return evalVBASingle(code);
}

console.log('--- Starting Enum Declaration Tests ---');

const enumCode = `
    Enum Color
        Red = 1
        Green
        Blue = 10
        Yellow
    End Enum
`;

const ev = evalVBA(enumCode);

// --- メンバ名による直接アクセス ---
{
    assert.strictEqual(ev.evalExpression('Red'), 1, 'Red = 1 (明示)');
    assert.strictEqual(ev.evalExpression('Green'), 2, 'Green = 2 (自動インクリメント)');
    assert.strictEqual(ev.evalExpression('Blue'), 10, 'Blue = 10 (明示)');
    assert.strictEqual(ev.evalExpression('Yellow'), 11, 'Yellow = 11 (自動インクリメント)');
    console.log('[PASS] メンバ名による直接アクセス');
}

// --- Enum 名経由のアクセス ---
{
    assert.strictEqual(ev.evalExpression('Color.Red'), 1, 'Color.Red = 1');
    assert.strictEqual(ev.evalExpression('Color.Green'), 2, 'Color.Green = 2');
    assert.strictEqual(ev.evalExpression('Color.Blue'), 10, 'Color.Blue = 10');
    assert.strictEqual(ev.evalExpression('Color.Yellow'), 11, 'Color.Yellow = 11');
    console.log('[PASS] Enum 名経由のアクセス');
}

// Bug CB: `Dim c As Color` — TypeName(c) が "Double" を返していた（Enum型変数のタイプ追跡が未対応）
{
    const code = `
Enum Priority
    Low = 1
    Medium = 5
    High = 10
End Enum
Function TestCB_TypeName() As String
    Dim p As Priority
    p = Medium
    TestCB_TypeName = TypeName(p) & "," & VarType(p)
End Function
Function TestCB_VarType() As Long
    Dim p As Priority
    p = High
    TestCB_VarType = VarType(p)
End Function
`;
    const ev2 = evalVBASingle(code);
    assert.strictEqual(ev2.callProcedure('TestCB_TypeName', []), 'Long,3', 'Bug CB: Dim p As Priority → TypeName(p)="Long", VarType(p)=3');
    assert.strictEqual(ev2.callProcedure('TestCB_VarType', []), 3, 'Bug CB: VarType(p)=3 (vbLong) for Enum-typed variable');
    console.log('[PASS] Bug CB: Enum 型変数の TypeName/VarType が "Long"/3 を返す');
}

console.log('\n✅ Enum 宣言: 全テスト通過');
