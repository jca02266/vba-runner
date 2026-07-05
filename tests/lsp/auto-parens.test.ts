import { assert } from '../../test-libs/test-runner';
import { autoParensEdit, getBlockEnd, needsBodyIndent } from '../../src/lsp/auto-parens';

// ─── autoParensEdit ─────────────────────────────────────────────────────────

function col(lineText: string): number | null {
    return autoParensEdit(lineText)?.insertCol ?? null;
}

// 括弧なし Sub → 挿入位置を返す
assert.strictEqual(col('Sub Foo'), 7, 'Sub Foo → insertCol:7');
assert.strictEqual(col('Function Bar'), 12, 'Function Bar → insertCol:12');
assert.strictEqual(col('Private Sub Calc'), 16, 'Private Sub Calc → insertCol:16');

// 戻り型付き Function（Bug 1 修正確認）
assert.strictEqual(col('Function GetValue As Long'), 17, 'Function GetValue As Long → insertCol:17');
assert.strictEqual(col('Private Function Calc As Double'), 21, 'Private Function Calc As Double → insertCol:21');

// 括弧あり → null
assert.strictEqual(autoParensEdit('Sub Foo()'), null, 'Sub Foo() → null');
assert.strictEqual(autoParensEdit('Function Bar()'), null, 'Function Bar() → null');
assert.strictEqual(autoParensEdit('Function GetValue() As Long'), null, 'Function GetValue() As Long → null');

// プロシージャ宣言でない行 → null
assert.strictEqual(autoParensEdit('Dim x As Long'), null, 'Dim line → null');
assert.strictEqual(autoParensEdit('    x = 1'), null, 'assignment → null');

console.log('[PASS] autoParensEdit');

// ─── getBlockEnd ─────────────────────────────────────────────────────────────

assert.strictEqual(getBlockEnd('Sub Foo()')?.insertKeyword, 'End Sub', 'Sub → End Sub');
assert.strictEqual(getBlockEnd('Function Bar()')?.insertKeyword, 'End Function', 'Function → End Function');
assert.strictEqual(getBlockEnd('Property Get X()')?.insertKeyword, 'End Property', 'Property Get → End Property');
assert.strictEqual(getBlockEnd('If x > 0 Then')?.insertKeyword, 'End If', 'If Then → End If');
assert.strictEqual(getBlockEnd('For i = 1 To 10')?.insertKeyword?.startsWith('Next'), true, 'For → Next ...');
assert.strictEqual(getBlockEnd('With obj')?.insertKeyword, 'End With', 'With → End With');
assert.strictEqual(getBlockEnd('Select Case x')?.insertKeyword, 'End Select', 'Select Case → End Select');
assert.strictEqual(getBlockEnd('Do While x > 0')?.insertKeyword, 'Loop', 'Do While → Loop');
assert.strictEqual(getBlockEnd('While x > 0')?.insertKeyword, 'Wend', 'While → Wend');
assert.strictEqual(getBlockEnd('Dim x As Long'), null, 'Dim → null');

console.log('[PASS] getBlockEnd');

// ─── needsBodyIndent ─────────────────────────────────────────────────────────

assert.strictEqual(needsBodyIndent('Else'), true, 'Else → true');
assert.strictEqual(needsBodyIndent('ElseIf x > 0 Then'), true, 'ElseIf → true');
assert.strictEqual(needsBodyIndent('Case 1'), true, 'Case → true');
assert.strictEqual(needsBodyIndent('Case Else'), true, 'Case Else → true');
assert.strictEqual(needsBodyIndent('Sub Foo()'), false, 'Sub → false');
assert.strictEqual(needsBodyIndent('x = 1'), false, 'assignment → false');

console.log('[PASS] needsBodyIndent');

console.log('\n✅ auto-parens: 全テスト通過');
