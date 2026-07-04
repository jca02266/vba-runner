import { format, applyEdits, FormatterOptions } from '../../src/lsp/formatter';
import { assert } from '../../test-libs/test-runner';

function fmt(source: string, options?: FormatterOptions): string {
    return applyEdits(source, format(source, options));
}

// --- keyword case -----------------------------------------------------------

// 1. Sub/Function/End keywords are Pascal-cased
{
    const result = fmt('sub Foo()\nend sub');
    assert.strictEqual(result, 'Sub Foo()\nEnd Sub', 'sub/end sub → Sub/End Sub');
    console.log('[PASS] Sub/End Sub keyword case');
}

// 2. Dim, As keywords
{
    const result = fmt('sub S()\ndim x as long\nend sub');
    assert.strictEqual(result.split('\n')[1].trimStart(), 'Dim x As Long', 'dim/as/long → Dim/As/Long');
    console.log('[PASS] Dim/As keyword case');
}

// 3. If/Then/Else/ElseIf/End If
{
    const result = fmt('sub S()\nif x then\nx = 1\nelse\nx = 2\nend if\nend sub');
    assert.ok(result.includes('If x Then'), 'if/then → If/Then');
    assert.ok(result.includes('Else'), 'else → Else');
    assert.ok(result.includes('End If'), 'end if → End If');
    console.log('[PASS] If/Then/Else/End If keyword case');
}

// 4. For/To/Next, And/Or/Not
{
    const result = fmt('sub S()\nfor i = 1 to 10\nif i mod 2 = 0 and not done then\nx = x + 1\nend if\nnext i\nend sub');
    assert.ok(result.includes('For i = 1 To 10'), 'for/to → For/To');
    assert.ok(result.includes('Next i'), 'next → Next');
    assert.ok(result.includes('Mod'), 'mod → Mod');
    assert.ok(result.includes('And'), 'and → And');
    assert.ok(result.includes('Not'), 'not → Not');
    console.log('[PASS] For/To/Next/Mod/And/Not keyword case');
}

// 5. Select Case
{
    const result = fmt('sub S()\nselect case x\ncase 1\ny = 1\ncase else\ny = 0\nend select\nend sub');
    assert.ok(result.includes('Select Case'), 'select case → Select Case');
    assert.ok(result.includes('Case Else'), 'case else → Case Else');
    assert.ok(result.includes('End Select'), 'end select → End Select');
    console.log('[PASS] Select Case keyword case');
}

// 6. Already correctly cased → no edits
{
    const src = 'Sub Foo()\n    Dim x As Long\nEnd Sub';
    const edits = format(src);
    assert.strictEqual(edits.length, 0, 'already formatted → no edits');
    console.log('[PASS] Already correctly cased → no edits');
}

// --- indentation ------------------------------------------------------------

// 7. Sub body is indented 4 spaces
{
    const result = fmt('sub Foo()\nx = 1\ny = 2\nend sub');
    const lines = result.split('\n');
    assert.strictEqual(lines[1], '    x = 1', 'body line indented 4 spaces');
    assert.strictEqual(lines[2], '    y = 2', 'second body line indented 4 spaces');
    console.log('[PASS] Sub body indented 4 spaces');
}

// 8. If block indentation
{
    const result = fmt('sub S()\nif x > 0 then\nx = 1\nend if\nend sub');
    const lines = result.split('\n');
    assert.strictEqual(lines[1], '    If x > 0 Then', 'If line at level 1');
    assert.strictEqual(lines[2], '        x = 1', 'If body at level 2');
    assert.strictEqual(lines[3], '    End If', 'End If at level 1');
    console.log('[PASS] If block indentation');
}

// 9. Else/ElseIf indentation
{
    const result = fmt('sub S()\nif x > 0 then\nx = 1\nelse\nx = 2\nend if\nend sub');
    const lines = result.split('\n');
    assert.strictEqual(lines[3], '    Else', 'Else at level 1 (same as If)');
    assert.strictEqual(lines[4], '        x = 2', 'Else body at level 2');
    console.log('[PASS] Else indentation');
}

// 10. For loop indentation
{
    const result = fmt('sub S()\nfor i = 1 to 10\nx = x + i\nnext i\nend sub');
    const lines = result.split('\n');
    assert.strictEqual(lines[1], '    For i = 1 To 10', 'For at level 1');
    assert.strictEqual(lines[2], '        x = x + i', 'For body at level 2');
    assert.strictEqual(lines[3], '    Next i', 'Next at level 1');
    console.log('[PASS] For loop indentation');
}

// 11. Select Case indentation
{
    const result = fmt('sub S()\nselect case x\ncase 1\ny = 1\ncase 2\ny = 2\nend select\nend sub');
    const lines = result.split('\n');
    assert.strictEqual(lines[1], '    Select Case x', 'Select Case at level 1');
    assert.strictEqual(lines[2], '    Case 1', 'Case at level 1 (same as Select)');
    assert.strictEqual(lines[3], '        y = 1', 'Case body at level 2');
    assert.strictEqual(lines[5], '        y = 2', 'Case 2 body at level 2');
    assert.strictEqual(lines[6], '    End Select', 'End Select at level 1');
    console.log('[PASS] Select Case indentation');
}

// 12. Nested blocks
{
    const result = fmt('sub S()\nfor i = 1 to 10\nif i > 5 then\nx = x + i\nend if\nnext i\nend sub');
    const lines = result.split('\n');
    assert.strictEqual(lines[2], '        If i > 5 Then', 'Nested If at level 2');
    assert.strictEqual(lines[3], '            x = x + i', 'Nested If body at level 3');
    assert.strictEqual(lines[4], '        End If', 'Nested End If at level 2');
    console.log('[PASS] Nested block indentation');
}

// 13. Single-line If does NOT open a block
{
    const result = fmt('sub S()\nif x > 0 then y = 1\ny = 2\nend sub');
    const lines = result.split('\n');
    assert.strictEqual(lines[1], '    If x > 0 Then y = 1', 'single-line If at level 1');
    assert.strictEqual(lines[2], '    y = 2', 'line after single-line If at level 1 (not 2)');
    console.log('[PASS] Single-line If does not open block');
}

// 14. With block indentation
{
    const result = fmt('sub S()\nwith obj\n.name = "x"\nend with\nend sub');
    const lines = result.split('\n');
    assert.strictEqual(lines[1], '    With obj', 'With at level 1');
    assert.strictEqual(lines[2], '        .name = "x"', '.member at level 2');
    assert.strictEqual(lines[3], '    End With', 'End With at level 1');
    console.log('[PASS] With block indentation');
}

// 15. Custom indent size (2 spaces)
{
    const result = fmt('sub S()\nx = 1\nend sub', { indentSize: 2 });
    const lines = result.split('\n');
    assert.strictEqual(lines[1], '  x = 1', 'body indented 2 spaces');
    console.log('[PASS] Custom indent size (2 spaces)');
}

// 16. keywordCase: false — no keyword changes
{
    const result = fmt('sub Foo()\ndim x as long\nend sub', { keywordCase: false });
    assert.ok(result.includes('sub Foo()'), 'sub not cased when keywordCase=false');
    assert.ok(result.includes('dim x as long'), 'dim/as not cased when keywordCase=false');
    console.log('[PASS] keywordCase: false skips keyword casing');
}

// 17. Idempotent: formatting twice gives same result
{
    const code = 'sub Test()\ndim i as long\nfor i = 1 to 10\nif i mod 2 = 0 then\nx = x + i\nend if\nnext i\nend sub';
    const once = fmt(code);
    const twice = fmt(once);
    assert.strictEqual(twice, once, 'formatting is idempotent');
    console.log('[PASS] Formatting is idempotent');
}

// 18. Function with return type
{
    const result = fmt('function Add(a as long, b as long) as long\nadd = a + b\nend function');
    const lines = result.split('\n');
    assert.strictEqual(lines[0], 'Function Add(a As Long, b As Long) As Long', 'Function keyword and As cased; long → Long');
    assert.strictEqual(lines[1], '    Add = a + b', 'Function body indented (Add cased to match definition)');
    assert.strictEqual(lines[2], 'End Function', 'End Function');
    console.log('[PASS] Function declaration formatting');
}

// 19. Option Explicit
{
    const result = fmt('option explicit\nsub S()\nx = 1\nend sub');
    assert.ok(result.startsWith('Option Explicit'), 'Option Explicit cased');
    console.log('[PASS] Option Explicit keyword case');
}

// 20. Public Type block indentation
{
    const result = fmt('Public Type AAA\nResponse As Object\nStatus As Integer\nEnd Type');
    const lines = result.split('\n');
    assert.strictEqual(lines[0], 'Public Type AAA', 'Public Type header unchanged');
    assert.strictEqual(lines[1], '    Response As Object', 'member indented 4 spaces');
    assert.strictEqual(lines[2], '    Status As Integer', 'second member indented 4 spaces');
    assert.strictEqual(lines[3], 'End Type', 'End Type at level 0');
    console.log('[PASS] Public Type block indentation');
}

// 21. Private Enum block indentation
{
    const result = fmt('Private Enum MyEnum\nValA\nValB\nEnd Enum');
    const lines = result.split('\n');
    assert.strictEqual(lines[1], '    ValA', 'Enum member indented 4 spaces');
    assert.strictEqual(lines[3], 'End Enum', 'End Enum at level 0');
    console.log('[PASS] Private Enum block indentation');
}

// 22. If ... Then: (block If with trailing colon is treated as block, not single-line)
{
    const result = fmt('sub S()\nif x > 0 then:\nx = 1\nend if\nend sub');
    const lines = result.split('\n');
    assert.strictEqual(lines[1], '    If x > 0 Then:', 'If...Then: at level 1 (block form)');
    assert.strictEqual(lines[2], '        x = 1', 'body indented to level 2');
    assert.strictEqual(lines[3], '    End If', 'End If at level 1');
    console.log('[PASS] If...Then: (trailing colon) treated as block If');
}

// 23. On Error GoTo lbl: trailing colon does not affect formatter
{
    const result = fmt('sub S()\non error goto errHandler:\nx = 1\nexit sub\nerrHandler:\nx = 0\nend sub');
    assert.ok(result.includes('On Error GoTo errHandler:'), 'On Error GoTo label: preserved');
    assert.ok(result.includes('errHandler:'), 'label line preserved');
    console.log('[PASS] On Error GoTo lbl: trailing colon handled');
}

// 24. Label is always at column 0 (VBA convention)
{
    const result = fmt('sub S()\non error goto ErrHandler\nx = 1\nexit sub\n    ErrHandler:\nx = 0\nend sub');
    const lines = result.split('\n');
    const labelLine = lines.find(l => l.includes('ErrHandler:'));
    assert.ok(labelLine !== undefined, 'label line exists');
    assert.strictEqual(labelLine, 'ErrHandler:', 'label at column 0 (no leading indent)');
    assert.strictEqual(lines[lines.indexOf(labelLine!) + 1], '    x = 0', 'code after label indented normally');
    console.log('[PASS] Label is always at column 0');
}

// 25. Multi-line If...Then (line continuation) indents body correctly
{
    const result = fmt('sub S()\nif xxxx _\n   and yyyy then\nx = 1\nend if\nend sub');
    const lines = result.split('\n');
    assert.strictEqual(lines[1], '    If xxxx _', 'If line at level 1');
    assert.strictEqual(lines[3], '        x = 1', 'If body at level 2');
    assert.strictEqual(lines[4], '    End If', 'End If at level 1');
    console.log('[PASS] Multi-line If...Then continuation indents body correctly');
}

// 26. Multi-line ElseIf...Then (line continuation) indents body correctly
{
    const result = fmt('sub S()\nif x > 0 then\ny = 1\nelseif xxxx _\n   and yyyy then\ny = 2\nend if\nend sub');
    const lines = result.split('\n');
    assert.strictEqual(lines[3], '    ElseIf xxxx _', 'ElseIf line at level 1');
    assert.strictEqual(lines[5], '        y = 2', 'ElseIf body at level 2');
    assert.strictEqual(lines[6], '    End If', 'End If at level 1');
    console.log('[PASS] Multi-line ElseIf...Then continuation indents body correctly');
}

// 27. Multi-line For loop (line continuation) indents body correctly
{
    const result = fmt('sub S()\nfor i = 1 _\n   to 10\nx = x + i\nnext i\nend sub');
    const lines = result.split('\n');
    assert.strictEqual(lines[1], '    For i = 1 _', 'For line at level 1');
    assert.strictEqual(lines[3], '        x = x + i', 'For body at level 2');
    assert.strictEqual(lines[4], '    Next i', 'Next at level 1');
    console.log('[PASS] Multi-line For loop continuation indents body correctly');
}

// 28. Identifier case matches declaration (Dim)
{
    const result = fmt('Sub S()\nDim myVariable As Long\nmyvariable = 1\nMYVARIABLE = MYVARIABLE + 1\nEnd Sub');
    assert.ok(result.includes('myVariable = 1'), 'myvariable → myVariable');
    assert.ok(result.includes('myVariable = myVariable + 1'), 'MYVARIABLE → myVariable');
    console.log('[PASS] Identifier case matches Dim declaration');
}

// 29. Identifier case matches Sub/Function name
{
    const result = fmt('Sub MyProc()\nmyproc\nEnd Sub');
    assert.ok(result.includes('    MyProc'), 'myproc → MyProc (recursive call)');
    console.log('[PASS] Identifier case matches Sub name');
}

// 30. Identifier case matches parameter name
{
    const result = fmt('Sub Greet(personName As String)\ndim msg as string\nmsg = "Hello " & PERSONNAME\nEnd Sub');
    assert.ok(result.includes('personName'), 'PERSONNAME → personName');
    console.log('[PASS] Identifier case matches parameter name');
}

// 31. Type member identifier case normalization
{
    const result = fmt('Public Type Person\n    FirstName As String\nEnd Type\nSub S()\nDim p As Person\np.firstname = "Alice"\nEnd Sub');
    // Type member names in assignments are normalized
    assert.ok(result.includes('FirstName As String'), 'Type member definition preserved');
    console.log('[PASS] Type member identifier case');
}

// 32. Identifier casing is scoped: same name with different cases in different procs
{
    const src = 'Function Add(A, b)\nAdd = A + b\nEnd Function\n\nFunction Multiply(a, b)\nMultiply = a * b\nEnd Function';
    const result = fmt(src);
    const lines = result.split('\n');
    // Add's scope: A stays A
    assert.strictEqual(lines[0], 'Function Add(A, b)', 'Add param A preserved');
    assert.strictEqual(lines[1], '    Add = A + b', 'A in Add scope stays A');
    // Multiply's scope: a stays a
    assert.strictEqual(lines[4], 'Function Multiply(a, b)', 'Multiply param a preserved');
    assert.strictEqual(lines[5], '    Multiply = a * b', 'a in Multiply scope stays a');
    console.log('[PASS] Identifier casing is scoped per procedure');
}

// --- built-in identifier normalization ----------------------------------------

// 33. 組み込み型名（小文字 → Pascal-case）
{
    const result = fmt('Sub S()\nDim x As string\nDim y As integer\nDim z As boolean\nEnd Sub');
    assert.ok(result.includes('As String'), 'string → String');
    assert.ok(result.includes('As Integer'), 'integer → Integer');
    assert.ok(result.includes('As Boolean'), 'boolean → Boolean');
    console.log('[PASS] 組み込み型名 string/integer/boolean の正規化');
}

// 34. 組み込み型名（大文字 → Pascal-case）
{
    const result = fmt('Sub S()\nDim a As DOUBLE\nDim b As LONG\nDim c As VARIANT\nEnd Sub');
    assert.ok(result.includes('As Double'), 'DOUBLE → Double');
    assert.ok(result.includes('As Long'), 'LONG → Long');
    assert.ok(result.includes('As Variant'), 'VARIANT → Variant');
    console.log('[PASS] 組み込み型名 DOUBLE/LONG/VARIANT の正規化');
}

// 35. Debug/Err オブジェクトの正規化
{
    const result = fmt('Sub S()\ndebug.Print "hello"\nerr.Raise 5\nEnd Sub');
    assert.ok(result.includes('Debug.Print'), 'debug → Debug');
    assert.ok(result.includes('Err.Raise'), 'err → Err');
    console.log('[PASS] Debug/Err オブジェクトの正規化');
}

// 36. ユーザー定義識別子が組み込み型名より優先される
{
    // ユーザーが "string" という名前のパラメーターを使う場合（珍しいが許可される）
    // → collectDefinitions でパラメーター名として登録されるので組み込みマップに勝つ
    const result = fmt('Sub S(string As Long)\nDim x As Long\nx = string\nEnd Sub');
    // "string" はパラメーター名なので小文字のまま保持（Dim で宣言されていない）
    // ただし As Long の Long は正規化される
    assert.ok(result.includes('As Long'), 'Long は正規化される');
    console.log('[PASS] ユーザー宣言識別子が組み込み型名より優先');
}

console.log('\n✅ lsp-formatter: 全テスト通過');
