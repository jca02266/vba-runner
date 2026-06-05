/**
 * §3.3.5.2: contextual keyword を For/ForEach/Enum/Type の識別子に使うテスト
 *
 * parseForStatement / parseForEachStatementBody / parseEnumDeclaration /
 * parseTypeDeclaration が isIdentifier() を使わず TokenType.Identifier のみで
 * チェックしていたため、contextual keyword を名前に使えなかった問題を修正するテスト。
 */
import { evalVBASingle, assert } from '../../test-libs/test-runner';

function run(code: string, name = 'T'): any {
    return evalVBASingle(code).callProcedure(name, []);
}

// --- 1. For ループのカウンター変数名に contextual keyword ---
{
    // Output (CONTEXTUAL_KW_FILE_MODE)
    assert.strictEqual(
        run(`Function T() As Long
Dim Output As Long
Dim s As Long
s = 0
For Output = 1 To 3
    s = s + Output
Next Output
T = s
End Function`),
        6, 'For Output = 1 To 3'
    );

    // Step (CONTEXTUAL_KW_STMT_ABSENT) — Step はループの step 指定にも使われるが変数名としても可
    assert.strictEqual(
        run(`Function T() As Long
Dim Step As Long
Dim s As Long
s = 0
For Step = 1 To 3
    s = s + Step
Next Step
T = s
End Function`),
        6, 'For Step = 1 To 3'
    );

    // Binary (CONTEXTUAL_KW_FILE_MODE)
    assert.strictEqual(
        run(`Function T() As Long
Dim Binary As Long
For Binary = 10 To 12
Next Binary
T = Binary
End Function`),
        13, 'For Binary = 10 To 12 (post-loop value)'
    );

    // Access (CONTEXTUAL_KW_FILE_ACCESS)
    assert.strictEqual(
        run(`Function T() As Long
Dim Access As Long
Dim s As Long
For Access = 1 To 4
    s = s + 1
Next Access
T = s
End Function`),
        4, 'For Access = 1 To 4'
    );

    // Class (CONTEXTUAL_KW_STRUCTURAL)
    assert.strictEqual(
        run(`Function T() As Long
Dim Class As Long
For Class = 5 To 7
Next Class
T = Class
End Function`),
        8, 'For Class = 5 To 7 (post-loop value)'
    );

    console.log('[PASS] For ループのカウンター変数名に contextual keyword');
}

// --- 2. For ループで Next の後ろにも contextual keyword ---
{
    // Next の後ろのラベルが contextual keyword でも正しく照合される
    assert.strictEqual(
        run(`Function T() As Long
Dim Output As Long
Dim s As Long
For Output = 1 To 5
    s = s + Output
Next Output
T = s
End Function`),
        15, 'Next Output (contextual keyword after Next)'
    );

    console.log('[PASS] Next の後ろの contextual keyword');
}

// --- 3. For Each のイテレーター変数名に contextual keyword ---
{
    assert.strictEqual(
        run(`Function T() As Long
Dim Text As Variant
Dim arr(2) As Long
arr(0) = 1: arr(1) = 2: arr(2) = 3
Dim s As Long
For Each Text In arr
    s = s + Text
Next Text
T = s
End Function`),
        6, 'For Each Text In arr'
    );

    assert.strictEqual(
        run(`Function T() As Long
Dim Output As Variant
Dim arr(1) As Long
arr(0) = 10: arr(1) = 20
Dim s As Long
For Each Output In arr
    s = s + Output
Next Output
T = s
End Function`),
        30, 'For Each Output In arr'
    );

    console.log('[PASS] For Each のイテレーター変数名に contextual keyword');
}

// --- 4. Enum 宣言の名前に contextual keyword ---
{
    // Enum の名前として contextual keyword を使う
    assert.strictEqual(
        run(`
Enum Output
    Val1 = 1
    Val2 = 2
End Enum

Function T() As Long
    T = Output.Val1 + Output.Val2
End Function`),
        3, 'Enum Output { Val1, Val2 }'
    );

    assert.strictEqual(
        run(`
Enum Access
    ReadOnly = 1
    ReadWrite = 3
End Enum

Function T() As Long
    T = Access.ReadWrite
End Function`),
        3, 'Enum Access { ReadOnly, ReadWrite }'
    );

    console.log('[PASS] Enum 宣言の名前に contextual keyword');
}

// --- 5. Type 宣言の名前に contextual keyword ---
{
    assert.strictEqual(
        run(`
Type Output
    X As Long
    Y As Long
End Type

Function T() As Long
    Dim o As Output
    o.X = 4
    o.Y = 5
    T = o.X + o.Y
End Function`),
        9, 'Type Output { X, Y }'
    );

    assert.strictEqual(
        run(`
Type Access
    Name As String
    Level As Long
End Type

Function T() As Long
    Dim a As Access
    a.Level = 7
    T = a.Level
End Function`),
        7, 'Type Access { Name, Level }'
    );

    console.log('[PASS] Type 宣言の名前に contextual keyword');
}

console.log('\n✅ contextual-kw-for-enum-type: 全テスト通過');
