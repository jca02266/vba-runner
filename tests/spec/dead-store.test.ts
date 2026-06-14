/**
 * デッドストア検出テスト — Phase 4D
 */
import { Lexer } from '../../src/engine/lexer';
import { Parser } from '../../src/engine/parser';
import { findDeadStores, DeadStore } from '../../src/engine/dead-store';
import { ProcedureDeclaration } from '../../src/engine/parser';
import { assert } from '../../test-libs/test-runner';

function analyze(code: string): DeadStore[] {
    const tokens = new Lexer(code).tokenize();
    const ast = new Parser(tokens).parse();
    const proc = ast.body.find(
        (s): s is ProcedureDeclaration => s.type === 'ProcedureDeclaration',
    );
    if (!proc) throw new Error('No ProcedureDeclaration found');
    return findDeadStores(proc);
}

function varNames(ds: DeadStore[]): string[] {
    return ds.map(d => d.varName).sort();
}

// ─── 基本ケース ──────────────────────────────────────────────────────────────

// 使われる変数 → デッドストアなし
{
    const ds = analyze(`
Sub Test()
    Dim x As Long
    x = 5
    Debug.Print x
End Sub
`);
    assert.strictEqual(ds.length, 0, '使用済み: デッドストアなし');
    console.log('[PASS] 使用済み変数はデッドストアなし');
}

// 代入後に使用されない
{
    const ds = analyze(`
Sub Test()
    Dim x As Long
    x = 5
End Sub
`);
    assert.strictEqual(varNames(ds).includes('x'), true, '未使用代入: x はデッドストア');
    console.log('[PASS] 代入後未使用: デッドストア検出');
}

// 2 度代入して最初の代入が死んでいる
{
    const ds = analyze(`
Sub Test()
    Dim x As Long
    x = 1
    x = 2
    Debug.Print x
End Sub
`);
    assert.strictEqual(varNames(ds).includes('x'), true, '上書き: 最初の x = 1 はデッドストア');
    console.log('[PASS] 上書き代入: 最初の代入をデッドストア検出');
}

// x = x + 1 の形式: RHS で使用するので最初の代入は生きている
{
    const ds = analyze(`
Sub Test()
    Dim x As Long
    x = 1
    x = x + 1
    Debug.Print x
End Sub
`);
    assert.strictEqual(ds.length, 0, 'x = x + 1: デッドストアなし');
    console.log('[PASS] x = x + 1: デッドストアなし');
}

// ─── 条件分岐 ────────────────────────────────────────────────────────────────

// If 分岐: 両方のパスで x が使用される
{
    const ds = analyze(`
Sub Test()
    Dim x As Long
    x = 0
    If flag Then
        Debug.Print x
    Else
        Debug.Print x
    End If
End Sub
`);
    assert.strictEqual(ds.length, 0, 'If 両パス使用: デッドストアなし');
    console.log('[PASS] If 両パス使用: デッドストアなし');
}

// If 分岐: 一方のパスでのみ使用 → 片方だけでも生きているのでデッドストアなし
{
    const ds = analyze(`
Sub Test()
    Dim x As Long
    x = 0
    If flag Then
        Debug.Print x
    End If
End Sub
`);
    assert.strictEqual(ds.length, 0, 'If 片パス使用: デッドストアなし（保守的）');
    console.log('[PASS] If 片パス使用: デッドストアなし（保守的解析）');
}

// ─── ループ ──────────────────────────────────────────────────────────────────

// For ループカウンターは使用される
{
    const ds = analyze(`
Sub Test()
    Dim i As Long
    For i = 1 To 10
        Debug.Print i
    Next i
End Sub
`);
    assert.strictEqual(ds.length, 0, 'For カウンター使用: デッドストアなし');
    console.log('[PASS] For ループカウンター: デッドストアなし');
}

// ループ内代入でループ後に使用されない
{
    const ds = analyze(`
Sub Test()
    Dim i As Long
    Dim s As Long
    For i = 1 To 5
        s = i * 2
    Next i
End Sub
`);
    // s はループ後使用されない → デッドストア
    assert.strictEqual(varNames(ds).includes('s'), true, 'ループ後未使用: s はデッドストア');
    console.log('[PASS] ループ内代入がループ後未使用: デッドストア検出');
}

// ─── 関数戻り値 ───────────────────────────────────────────────────────────────

// 関数名への代入は戻り値なのでデッドストアではない
{
    const ds = analyze(`
Function Calc() As Long
    Dim x As Long
    x = 5
    Calc = x * 2
End Function
`);
    assert.strictEqual(varNames(ds).includes('calc'), false, '関数戻り値: calc はデッドストアでない');
    console.log('[PASS] 関数戻り値代入: デッドストアなし');
}

// 関数内のローカル変数で未使用
{
    const ds = analyze(`
Function Calc() As Long
    Dim tmp As Long
    tmp = 42
    Calc = 100
End Function
`);
    assert.strictEqual(varNames(ds).includes('tmp'), true, '関数内未使用ローカル: tmp はデッドストア');
    console.log('[PASS] 関数内未使用ローカル: デッドストア検出');
}

// ─── パラメーター ─────────────────────────────────────────────────────────────

// ByRef パラメーターへの代入は呼び元に影響するので除外
{
    const ds = analyze(`
Sub Test(ByRef result As Long)
    result = 42
End Sub
`);
    assert.strictEqual(varNames(ds).includes('result'), false, 'ByRef パラメーター: デッドストアでない');
    console.log('[PASS] ByRef パラメーター: デッドストアなし');
}

// ByVal パラメーターは関数内の代入が未使用なら Dead Store
// （パラメーター自体はスキップするので報告されない）
{
    const ds = analyze(`
Sub Test(ByVal n As Long)
    n = n + 1
End Sub
`);
    // パラメーターへの代入はスキップ
    assert.strictEqual(varNames(ds).includes('n'), false, 'パラメーター: デッドストア報告なし');
    console.log('[PASS] パラメーターへの代入: デッドストア報告なし（パラメーターはスキップ）');
}

// ─── Select Case ─────────────────────────────────────────────────────────────
{
    const ds = analyze(`
Sub Test()
    Dim result As Long
    Select Case x
        Case 1
            result = 10
        Case 2
            result = 20
        Case Else
            result = 0
    End Select
    Debug.Print result
End Sub
`);
    assert.strictEqual(ds.length, 0, 'Select Case: 全ケースで使用 → デッドストアなし');
    console.log('[PASS] Select Case: デッドストアなし');
}

// ─── モジュールレベル変数（クラスフィールド） ─────────────────────────────────────

// Init でクラスフィールドに代入 → Property Get で参照 → デッドストアなし
{
    // findDeadStores は ProcedureDeclaration 単体を受け取るため、
    // Init 内の m_row 代入が「プロシージャ内で参照されない = デッドストア」と
    // 誤判定されないことを確認する。
    const code = `
Private m_row As Long
Private m_value As String

Public Sub Init(ByVal r As Long)
    m_row   = r
    m_value = ""
End Sub

Public Property Get Row() As Long
    Row = m_row
End Property
`;
    const tokens = new Lexer(code).tokenize();
    const ast = new Parser(tokens).parse();
    const initProc = ast.body.find(
        (s): s is ProcedureDeclaration =>
            s.type === 'ProcedureDeclaration' && (s as ProcedureDeclaration).name.name.toLowerCase() === 'init',
    );
    if (!initProc) throw new Error('Init not found');
    const ds = findDeadStores(initProc);
    assert.strictEqual(ds.length, 0, 'クラスフィールドへの代入はデッドストアなし');
    console.log('[PASS] クラスフィールド代入: デッドストアなし');
}

console.log('\n✅ デッドストア検出: 全テスト通過');
