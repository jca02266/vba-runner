/**
 * vba-lint.ts のテスト
 * VBA 固有 Diagnostics 警告 (VBA001〜VBA009) の検出を確認する
 */
import { Lexer } from '../../src/engine/lexer';
import { Parser } from '../../src/engine/parser';
import { lintProgram, LintDiagnostic } from '../../src/engine/vba-lint';
import { assert } from '../../test-libs/test-runner';

function lint(code: string): LintDiagnostic[] {
    const tokens = new Lexer(code).tokenize();
    const ast = new Parser(tokens).parse();
    return lintProgram(ast);
}

// ─── VBA001: Dim 複数宣言で途中の変数に型指定なし ────────────────────────────
{
    const diags = lint('Sub Test()\n    Dim a, b As Long\nEnd Sub');
    const d = diags.filter(d => d.code === 'VBA001');
    assert.strictEqual(d.length, 1, 'VBA001: 1 件検出');
    assert.strictEqual(d[0].message.includes('a'), true, 'VBA001: 変数名 a を含む');
    assert.strictEqual(d[0].severity, 2, 'VBA001: severity = Warning');
    console.log('[PASS] VBA001: Dim 複数宣言 Variant トラップ');
}

// ─── VBA001: 全変数に型指定あり → 警告なし ──────────────────────────────────
{
    const diags = lint('Sub Test()\n    Dim a As Long, b As Long\nEnd Sub');
    const d = diags.filter(d => d.code === 'VBA001');
    assert.strictEqual(d.length, 0, 'VBA001: 型指定あり → 警告なし');
    console.log('[PASS] VBA001: 型指定あり → 警告なし');
}

// ─── VBA002: Integer 型変数 ───────────────────────────────────────────────────
{
    const diags = lint('Sub Test()\n    Dim n As Integer\nEnd Sub');
    const d = diags.filter(d => d.code === 'VBA002');
    assert.strictEqual(d.length, 1, 'VBA002: 1 件検出');
    assert.strictEqual(d[0].message.includes('n'), true, 'VBA002: 変数名 n を含む');
    assert.strictEqual(d[0].severity, 3, 'VBA002: severity = Information');
    console.log('[PASS] VBA002: Integer 型変数');
}

// ─── VBA002: Integer 型パラメーター ─────────────────────────────────────────
{
    const diags = lint('Sub Test(ByVal n As Integer)\nEnd Sub');
    const d = diags.filter(d => d.code === 'VBA002');
    assert.strictEqual(d.length, 1, 'VBA002: パラメーター Integer 1 件検出');
    assert.strictEqual(d[0].message.includes('n'), true, 'VBA002: パラメーター名 n を含む');
    console.log('[PASS] VBA002: Integer 型パラメーター');
}

// ─── VBA002: Long 型 → 警告なし ─────────────────────────────────────────────
{
    const diags = lint('Sub Test()\n    Dim n As Long\nEnd Sub');
    const d = diags.filter(d => d.code === 'VBA002');
    assert.strictEqual(d.length, 0, 'VBA002: Long → 警告なし');
    console.log('[PASS] VBA002: Long 型 → 警告なし');
}

// ─── VBA003: パラメーターに ByVal/ByRef なし ─────────────────────────────────
{
    const diags = lint('Sub Test(x As Long)\nEnd Sub');
    const d = diags.filter(d => d.code === 'VBA003');
    assert.strictEqual(d.length, 1, 'VBA003: 1 件検出');
    assert.strictEqual(d[0].message.includes('x'), true, 'VBA003: パラメーター名 x を含む');
    assert.strictEqual(d[0].severity, 2, 'VBA003: severity = Warning');
    console.log('[PASS] VBA003: ByVal/ByRef なし');
}

// ─── VBA003: ByVal 明示 → 警告なし ──────────────────────────────────────────
{
    const diags = lint('Sub Test(ByVal x As Long)\nEnd Sub');
    const d = diags.filter(d => d.code === 'VBA003');
    assert.strictEqual(d.length, 0, 'VBA003: ByVal 明示 → 警告なし');
    console.log('[PASS] VBA003: ByVal 明示 → 警告なし');
}

// ─── VBA003: ByRef 明示 → 警告なし ──────────────────────────────────────────
{
    const diags = lint('Sub Test(ByRef x As Long)\nEnd Sub');
    const d = diags.filter(d => d.code === 'VBA003');
    assert.strictEqual(d.length, 0, 'VBA003: ByRef 明示 → 警告なし');
    console.log('[PASS] VBA003: ByRef 明示 → 警告なし');
}

// ─── VBA004: While...Wend ────────────────────────────────────────────────────
{
    const diags = lint('Sub Test()\n    While True\n    Wend\nEnd Sub');
    const d = diags.filter(d => d.code === 'VBA004');
    assert.strictEqual(d.length, 1, 'VBA004: 1 件検出');
    assert.strictEqual(d[0].severity, 3, 'VBA004: severity = Information');
    console.log('[PASS] VBA004: While...Wend');
}

// ─── VBA005: Select Case に Case Else なし ───────────────────────────────────
{
    const code = [
        'Sub Test()',
        '    Dim x As Long',
        '    x = 1',
        '    Select Case x',
        '        Case 1',
        '            x = 2',
        '    End Select',
        'End Sub',
    ].join('\n');
    const diags = lint(code);
    const d = diags.filter(d => d.code === 'VBA005');
    assert.strictEqual(d.length, 1, 'VBA005: 1 件検出');
    assert.strictEqual(d[0].severity, 3, 'VBA005: severity = Information');
    console.log('[PASS] VBA005: Select Case without Case Else');
}

// ─── VBA005: Case Else あり → 警告なし ──────────────────────────────────────
{
    const code = [
        'Sub Test()',
        '    Dim x As Long',
        '    x = 1',
        '    Select Case x',
        '        Case 1',
        '            x = 2',
        '        Case Else',
        '            x = 0',
        '    End Select',
        'End Sub',
    ].join('\n');
    const diags = lint(code);
    const d = diags.filter(d => d.code === 'VBA005');
    assert.strictEqual(d.length, 0, 'VBA005: Case Else あり → 警告なし');
    console.log('[PASS] VBA005: Case Else あり → 警告なし');
}

// ─── VBA006: Sheets(1) 数値インデックス ─────────────────────────────────────
{
    const diags = lint('Sub Test()\n    Sheets(1).Activate\nEnd Sub');
    const d = diags.filter(d => d.code === 'VBA006');
    assert.strictEqual(d.length, 1, 'VBA006: Sheets(1) 検出');
    assert.strictEqual(d[0].severity, 2, 'VBA006: severity = Warning');
    console.log('[PASS] VBA006: Sheets(1) 数値インデックス');
}

// ─── VBA006: Worksheets(1) も検出 ───────────────────────────────────────────
{
    const diags = lint('Sub Test()\n    Worksheets(1).Activate\nEnd Sub');
    const d = diags.filter(d => d.code === 'VBA006');
    assert.strictEqual(d.length, 1, 'VBA006: Worksheets(1) 検出');
    console.log('[PASS] VBA006: Worksheets(1) 数値インデックス');
}

// ─── VBA006: Sheets("Sheet1") 文字列 → 警告なし ─────────────────────────────
{
    const diags = lint('Sub Test()\n    Sheets("Sheet1").Activate\nEnd Sub');
    const d = diags.filter(d => d.code === 'VBA006');
    assert.strictEqual(d.length, 0, 'VBA006: 文字列インデックス → 警告なし');
    console.log('[PASS] VBA006: 文字列インデックス → 警告なし');
}

// ─── VBA007: ActiveSheet 参照 ────────────────────────────────────────────────
{
    const diags = lint('Sub Test()\n    ActiveSheet.Activate\nEnd Sub');
    const d = diags.filter(d => d.code === 'VBA007');
    assert.strictEqual(d.length, 1, 'VBA007: ActiveSheet 検出');
    assert.strictEqual(d[0].severity, 3, 'VBA007: severity = Information');
    console.log('[PASS] VBA007: ActiveSheet 参照');
}

// ─── VBA007: ActiveWorkbook 参照 ─────────────────────────────────────────────
{
    const diags = lint('Sub Test()\n    Dim wb\n    Set wb = ActiveWorkbook\nEnd Sub');
    const d = diags.filter(d => d.code === 'VBA007');
    assert.strictEqual(d.length, 1, 'VBA007: ActiveWorkbook 検出');
    console.log('[PASS] VBA007: ActiveWorkbook 参照');
}

// ─── VBA008: GoTo（非エラーハンドラー） ─────────────────────────────────────
{
    const code = [
        'Sub Test()',
        '    GoTo MyLabel',
        '    MyLabel:',
        'End Sub',
    ].join('\n');
    const diags = lint(code);
    const d = diags.filter(d => d.code === 'VBA008');
    assert.strictEqual(d.length, 1, 'VBA008: GoTo 検出');
    assert.strictEqual(d[0].severity, 2, 'VBA008: severity = Warning');
    console.log('[PASS] VBA008: GoTo 非エラーハンドラー');
}

// ─── VBA008: GoTo ErrHandler → 除外 ─────────────────────────────────────────
{
    const code = [
        'Sub Test()',
        '    On Error GoTo ErrHandler',
        '    GoTo ErrHandler',
        '    Exit Sub',
        'ErrHandler:',
        'End Sub',
    ].join('\n');
    const diags = lint(code);
    const d = diags.filter(d => d.code === 'VBA008');
    assert.strictEqual(d.length, 0, 'VBA008: err含むラベル → 除外');
    console.log('[PASS] VBA008: エラーハンドラーラベル → 除外');
}

// ─── 複合: 複数ルールの同時検出 ─────────────────────────────────────────────
{
    const code = [
        'Sub Process(n As Integer)',
        '    Dim a, b As Long',
        '    While True',
        '    Wend',
        'End Sub',
    ].join('\n');
    const diags = lint(code);
    const codes = diags.map(d => d.code);
    assert.strictEqual(codes.includes('VBA001'), true, '複合: VBA001 あり');
    assert.strictEqual(codes.includes('VBA002'), true, '複合: VBA002 あり');
    assert.strictEqual(codes.includes('VBA003'), true, '複合: VBA003 あり');
    assert.strictEqual(codes.includes('VBA004'), true, '複合: VBA004 あり');
    console.log('[PASS] 複合: 複数ルール同時検出');
}

// ─── VBA009: デッドストア ────────────────────────────────────────────────────

// 代入後に未使用 → 検出
{
    const diags = lint([
        'Sub Test()',
        '    Dim x As Long',
        '    x = 42',
        'End Sub',
    ].join('\n'));
    const d = diags.filter(d => d.code === 'VBA009');
    assert.strictEqual(d.length >= 1, true, 'VBA009: x = 42 をデッドストア検出');
    assert.strictEqual(d[0].message.includes('x'), true, 'VBA009: 変数名 x を含む');
    assert.strictEqual(d[0].severity, 2, 'VBA009: severity = Warning');
    console.log('[PASS] VBA009: 代入後未使用のデッドストア検出');
}

// 上書き代入 → 最初の代入を検出
{
    const diags = lint([
        'Sub Test()',
        '    Dim x As Long',
        '    x = 1',
        '    x = 2',
        '    Debug.Print x',
        'End Sub',
    ].join('\n'));
    const d = diags.filter(d => d.code === 'VBA009');
    assert.strictEqual(d.length >= 1, true, 'VBA009: x = 1（上書き）をデッドストア検出');
    console.log('[PASS] VBA009: 上書き代入のデッドストア検出');
}

// 使用後に代入 → 検出なし
{
    const diags = lint([
        'Sub Test()',
        '    Dim x As Long',
        '    x = 5',
        '    Debug.Print x',
        'End Sub',
    ].join('\n'));
    const d = diags.filter(d => d.code === 'VBA009');
    assert.strictEqual(d.length, 0, 'VBA009: 使用済み変数 → 検出なし');
    console.log('[PASS] VBA009: 使用済み変数はデッドストアなし');
}

// 関数戻り値への代入 → 検出なし
{
    const diags = lint([
        'Function Calc() As Long',
        '    Calc = 42',
        'End Function',
    ].join('\n'));
    const d = diags.filter(d => d.code === 'VBA009');
    assert.strictEqual(d.length, 0, 'VBA009: 関数戻り値 → 検出なし');
    console.log('[PASS] VBA009: 関数戻り値代入はデッドストアなし');
}

// ByRef パラメーターへの代入 → 検出なし
{
    const diags = lint([
        'Sub Test(ByRef result As Long)',
        '    result = 99',
        'End Sub',
    ].join('\n'));
    const d = diags.filter(d => d.code === 'VBA009');
    assert.strictEqual(d.length, 0, 'VBA009: ByRef パラメーター → 検出なし');
    console.log('[PASS] VBA009: ByRef パラメーター代入はデッドストアなし');
}

// ─── VBA010: 到達不能コード ───────────────────────────────────────────────────

// Exit Sub の後ろのコード → 検出
{
    const code = [
        'Sub Test()',
        '    Exit Sub',
        '    Dim x As Long',
        '    x = 1',
        'End Sub',
    ].join('\n');
    const diags = lint(code);
    const d = diags.filter(d => d.code === 'VBA010');
    assert.strictEqual(d.length >= 1, true, 'VBA010: Exit Sub 後のコード検出');
    assert.strictEqual(d[0].severity, 3, 'VBA010: severity = Information');
    console.log('[PASS] VBA010: Exit Sub 後の到達不能コード');
}

// GoTo でスキップされたコード → 検出
{
    const code = [
        'Sub Test()',
        '    GoTo Done',
        '    Dim x As Long',
        'Done:',
        'End Sub',
    ].join('\n');
    const diags = lint(code);
    const d = diags.filter(d => d.code === 'VBA010');
    assert.strictEqual(d.length >= 1, true, 'VBA010: GoTo スキップ後のコード検出');
    console.log('[PASS] VBA010: GoTo スキップ後の到達不能コード');
}

// 到達可能なコードのみ → 警告なし
{
    const code = [
        'Sub Test()',
        '    Dim x As Long',
        '    x = 1',
        'End Sub',
    ].join('\n');
    const diags = lint(code);
    const d = diags.filter(d => d.code === 'VBA010');
    assert.strictEqual(d.length, 0, 'VBA010: 到達可能コード → 警告なし');
    console.log('[PASS] VBA010: 到達可能コード → 警告なし');
}

console.log('\n✅ VBA Lint: 全テスト通過');
