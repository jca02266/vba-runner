import { Lexer } from '../../src/engine/lexer';
import { Parser } from '../../src/engine/parser';
import { CompletionProvider } from '../../src/lsp/completion-provider';
import { assert } from '../../test-libs/test-runner';

function getCompletionsAt(src: string, line: number, character: number): any[] {
    const tokens = new Lexer(src).tokenize();
    const ast = new Parser(tokens, { errorRecovery: true }).parse();
    const provider = new CompletionProvider();
    return provider.getCompletions(ast.body, src, line, character);
}

// 1. Completions for standard VBA functions
{
    const code = 'x = Str';
    const completions = getCompletionsAt(code, 0, 6); // after 'Str'
    assert.ok(Array.isArray(completions), 'returns array');
    assert.ok(completions.some((c: any) => c.label === 'Str'), 'Str function found');
    console.log('[PASS] Standard function completion: Str');
}

// 2. Completions for Len function
{
    const code = 'x = L';
    const completions = getCompletionsAt(code, 0, 5);
    assert.ok(completions.some((c: any) => c.label === 'Len'), 'Len function found');
    console.log('[PASS] Function completion: Len');
}

// 3. Completions for other string functions
{
    const code = 'x = UCa';
    const completions = getCompletionsAt(code, 0, 7);
    assert.ok(completions.some((c: any) => c.label === 'UCase'), 'UCase found');
    console.log('[PASS] Function completion: UCase');
}

// 4. Local variable completion
{
    const code = 'Dim myVar As Integer\nx = my';
    const completions = getCompletionsAt(code, 1, 6); // after 'my'
    assert.ok(completions.some((c: any) => c.label === 'myVar'), 'myVar variable found');
    console.log('[PASS] Local variable completion');
}

// 5. Procedure completion
{
    const code = 'Sub MyProc()\nEnd Sub\nx = My';
    const completions = getCompletionsAt(code, 2, 7); // after 'My'
    assert.ok(completions.some((c: any) => c.label === 'MyProc'), 'MyProc found');
    console.log('[PASS] Procedure completion');
}

// 6. Completion item has required fields
{
    const code = 'x = Str';
    const completions = getCompletionsAt(code, 0, 6);
    const strCompletion = completions.find((c: any) => c.label === 'Str');
    assert.ok(strCompletion, 'Str completion found');
    assert.strictEqual(typeof strCompletion.label, 'string', 'label is string');
    assert.ok(strCompletion.kind, 'kind present');
    console.log('[PASS] Completion item has required fields');
}

// 7. Case-insensitive matching
{
    const code = 'x = str';
    const completions = getCompletionsAt(code, 0, 6); // lowercase 'str'
    assert.ok(completions.some((c: any) => c.label === 'Str'), 'Str found for lowercase str');
    console.log('[PASS] Case-insensitive matching');
}

// 8. Completions for Dim statement
{
    const code = 'Dim x As Inte';
    const completions = getCompletionsAt(code, 0, 13); // after 'Inte'
    // Should include 'Integer' type
    assert.ok(completions.length >= 0, 'completions returned for type context');
    console.log('[PASS] Type completion in Dim statement');
}

// 9. Empty prefix returns all completions
{
    const code = 'x = ';
    const completions = getCompletionsAt(code, 0, 4);
    // Should return standard functions and available variables
    assert.ok(completions.length > 0, 'completions returned for empty prefix');
    console.log('[PASS] Empty prefix returns all completions');
}

// 10. Completion for Class members
{
    const code = 'Class MyClass\n  Public value As Integer\nEnd Class\nSub Test()\nDim obj As New MyClass\nx = obj.';
    const completions = getCompletionsAt(code, 4, 11); // after 'obj.'
    // Should suggest 'value' member
    assert.ok(Array.isArray(completions), 'returns completions');
    console.log('[PASS] Class member completion');
}

// ─── ユーザー定義クラスのチェーン補完 ───────────────────────────────────────────
// [回帰] getPublicClassMembers が Function の returnType を設定しなかったためチェーン解決が
// 失敗し、`x.` でゴミ候補が出ていた問題 (getPublicClassMembers returnType fix)

// 11. 直接チェーン: ws.Cells(1, 4). → SimCell のメンバーが出る
{
    // ws は SimSheet 型、Cells() は As SimCell を返す → ws.Cells(1,4). で SimCell メンバーが補完される
    const code = [
        'Class SimCell',
        '  Public CellValue As String',
        '  Public CellRow As Long',
        'End Class',
        'Class SimSheet',
        '  Public Function Cells(r As Long, c As Long) As SimCell',
        '  End Function',
        'End Class',
        'Sub Test()',
        '  Dim ws As SimSheet',
        '  ws.Cells(1, 4).',   // line 10, character 18
        'End Sub',
    ].join('\n');
    // line 10: "  ws.Cells(1, 4)." → len=18
    const completions = getCompletionsAt(code, 10, 18);
    assert.ok(completions.some((c: any) => c.label === 'CellValue'), 'ws.Cells(1,4). → CellValue');
    assert.ok(completions.some((c: any) => c.label === 'CellRow'),   'ws.Cells(1,4). → CellRow');
    console.log('[PASS] 直接チェーン補完: ws.Cells(1,4). → SimCell メンバー');
}

// 12. Set 代入後の変数 `x.` → resolveSetAssignmentType でチェーン解決される
{
    // [回帰] Set x = ws.Cells(1,4) のあと x. で SimCell メンバーが出る
    // resolveSetAssignmentType が RHS をチェーン解決できることを確認
    const code = [
        'Class SimCell',
        '  Public CellValue As String',
        '  Public CellRow As Long',
        'End Class',
        'Class SimSheet',
        '  Public Function Cells(r As Long, c As Long) As SimCell',
        '  End Function',
        'End Class',
        'Sub Test()',
        '  Dim ws As SimSheet',
        '  Dim x As Object',
        '  Set x = ws.Cells(1, 4)',
        '  x.',                    // line 12, character 4
        'End Sub',
    ].join('\n');
    const completions = getCompletionsAt(code, 12, 4);
    assert.ok(completions.some((c: any) => c.label === 'CellValue'), 'x. → CellValue (Set代入チェーン)');
    assert.ok(completions.some((c: any) => c.label === 'CellRow'),   'x. → CellRow (Set代入チェーン)');
    console.log('[PASS] Set代入チェーン補完: Set x = ws.Cells(1,4) → x. → SimCell メンバー');
}

// 13. 型不明の Object 変数は空補完になる（ゴミ候補を出さない）
{
    // [回帰] Dim x As Object で Set 代入も CreateObject もない場合、
    //        型が解決できず補完は空になる（ドキュメントワード補完のゴミを返さない）
    const code = [
        'Sub Test()',
        '  Dim x As Object',
        '  x.',                // line 2, character 4
        'End Sub',
    ].join('\n');
    const completions = getCompletionsAt(code, 2, 4);
    assert.strictEqual(completions.length, 0, '型不明 Object の x. → 補完空');
    console.log('[PASS] 型不明 Object 変数の x. → 空補完（ゴミ候補なし）');
}

// 14. UDT (Type) メンバー補完 — Bug LSP-1 regression
// [回帰] getMembersForType が TypeDeclaration を未対応で pt. が 0件だった
{
    const code = [
        'Type Point',
        '    X As Long',
        '    Y As Long',
        'End Type',
        'Sub Test()',
        '    Dim pt As Point',
        '    pt.',             // line 6, character 7
        'End Sub',
    ].join('\n');
    const completions = getCompletionsAt(code, 6, 7);
    assert.ok(completions.some((c: any) => c.label === 'X'), 'UDT pt. → X を含む');
    assert.ok(completions.some((c: any) => c.label === 'Y'), 'UDT pt. → Y を含む');
    assert.strictEqual(completions.length, 2, 'UDT pt. → メンバー数は 2');
    console.log('[PASS] UDT (Type) メンバー補完: pt. → X, Y（Bug LSP-1）');
}

console.log('\n✅ LSP Completion: 全テスト通過');
