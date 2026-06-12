/**
 * §3.3.5.2: contextual keywords as label names in GoTo / GoSub / On...GoTo / Resume / RaiseEvent
 *
 * BNF:
 *   identifier-statement-label = IDENTIFIER
 *   goto-statement   = "GoTo" statement-label
 *   gosub-statement  = "GoSub" statement-label
 *   on-goto-statement = "On" expression "GoTo" statement-label-list
 *   on-gosub-statement = "On" expression "GoSub" statement-label-list
 *   resume-statement = "Resume" [statement-label]
 *   event-handler-name = IDENTIFIER       (Event declaration)
 *   RaiseEvent IDENTIFIER                 (RaiseEvent statement)
 *
 * Contextual keywords (Step, Error, Text, Binary, Compare, Access, Output,
 * Explicit, Lib, Alias, Module, Class, Property, Collection …) are IDENTIFIER
 * per §3.3.5.2 and MUST be accepted in all these positions.
 */
import { evalVBASingle, assert } from '../../test-libs/test-runner';

function parse(code: string): void {
    evalVBASingle(code, {
        onPrint: () => {},
        afterParse: (ast: any) => {
            if (ast.diagnostics?.length) {
                throw new Error(ast.diagnostics.map((d: any) => d.message).join('; '));
            }
        },
    });
}

function run(code: string, name: string): any {
    return evalVBASingle(code, { onPrint: () => {} }).callProcedure(name, []);
}

// ── GoTo <contextual-keyword> ─────────────────────────────────────────────────

{
    const keywords = [
        'Step', 'Error', 'Text', 'Binary', 'Compare', 'Access',
        'Output', 'Explicit', 'Lib', 'Alias', 'Module', 'Class',
        'Property', 'Collection', 'Append', 'Random', 'Read',
    ];
    for (const kw of keywords) {
        parse(`Sub S()\nGoTo ${kw}\nExit Sub\n${kw}:\nEnd Sub`);
    }
    console.log('[PASS] GoTo <contextual keyword>: 17件');
}

// ── GoSub <contextual-keyword> ────────────────────────────────────────────────

{
    // 動作確認: GoSub Step → Step: → Return
    const result = run(
        `Function F() As Long\nGoSub Step\nF = 1\nExit Function\nStep:\nReturn\nEnd Function`,
        'F',
    );
    assert.strictEqual(result, 1, 'GoSub Step executes and returns');

    const result2 = run(
        `Function F() As Long\nGoSub Error\nF = 2\nExit Function\nError:\nReturn\nEnd Function`,
        'F',
    );
    assert.strictEqual(result2, 2, 'GoSub Error executes and returns');
    console.log('[PASS] GoSub <contextual keyword>');
}

// ── On Error GoTo <contextual-keyword> ───────────────────────────────────────

{
    const result = run(
        `Function F() As Long\nOn Error GoTo Step\nF = 99\nExit Function\nStep:\nF = -1\nEnd Function`,
        'F',
    );
    assert.strictEqual(result, 99, 'On Error GoTo Step (no error path)');

    const result2 = run(
        `Function F() As Long\nOn Error GoTo Error\nF = 88\nExit Function\nError:\nF = -1\nEnd Function`,
        'F',
    );
    assert.strictEqual(result2, 88, 'On Error GoTo Error (no error path)');
    console.log('[PASS] On Error GoTo <contextual keyword>');
}

// ── On n GoTo <contextual-keyword-list> ──────────────────────────────────────

{
    parse(`Sub S()\nDim n As Long\nn = 1\nOn n GoTo Step\nExit Sub\nStep:\nEnd Sub`);
    parse(`Sub S()\nDim n As Long\nn = 2\nOn n GoTo Step, Error\nExit Sub\nStep:\nError:\nEnd Sub`);
    parse(`Sub S()\nDim n As Long\nn = 1\nOn n GoSub Step\nExit Sub\nStep:\nReturn\nEnd Sub`);
    console.log('[PASS] On n GoTo/GoSub <contextual keyword list>');
}

// ── Resume <contextual-keyword> ───────────────────────────────────────────────

{
    parse(`Sub S()\nOn Error GoTo ErrH\nError 1\nExit Sub\nErrH:\nResume Step\nStep:\nEnd Sub`);
    parse(`Sub S()\nOn Error GoTo ErrH\nError 1\nExit Sub\nErrH:\nResume Error\nError:\nEnd Sub`);
    console.log('[PASS] Resume <contextual keyword>');
}

// ── Event <contextual-keyword> / RaiseEvent <contextual-keyword> ──────────────

{
    parse(`Event Step()\nSub S()\nRaiseEvent Step\nEnd Sub`);
    parse(`Event Error(x As Long)\nSub S()\nRaiseEvent Error(1)\nEnd Sub`);
    parse(`Event Text()\nSub S()\nRaiseEvent Text\nEnd Sub`);
    console.log('[PASS] Event/RaiseEvent <contextual keyword>');
}

// ── Dim x As <contextual-keyword> (型名として) ───────────────────────────────

{
    // 型名としての contextual keyword（ユーザー定義型やクラス名として有効）
    // パース段階のエラーがないことを確認する
    parse(`Sub S()\nDim x As Error\nEnd Sub`);
    parse(`Sub S()\nDim x As Collection\nEnd Sub`);
    parse(`Sub S()\nDim x As Class\nEnd Sub`);
    parse(`Sub S()\nDim x As Text\nEnd Sub`);
    parse(`Sub S()\nDim x As Step\nEnd Sub`);
    parse(`Sub S()\nDim x As Property\nEnd Sub`);
    console.log('[PASS] Dim x As <contextual keyword> (型名)');
}

// ── Enum 名 / メンバー名 ─────────────────────────────────────────────────────

{
    parse(`Enum Step\n  Val1 = 1\nEnd Enum`);
    parse(`Enum E\n  Step = 1\n  Error = 2\n  Text = 3\nEnd Enum`);
    console.log('[PASS] Enum 名・メンバー名に contextual keyword');
}

// ── UDT メンバー名 ────────────────────────────────────────────────────────────

{
    parse(`Type T\n  Step As Long\n  Error As Long\n  Text As String\nEnd Type`);
    console.log('[PASS] UDT メンバー名に contextual keyword');
}

// ── Class <contextual-keyword> ────────────────────────────────────────────────

{
    parse(`Class Step\nEnd Class`);
    parse(`Class Error\nEnd Class`);
    parse(`Class Text\nEnd Class`);
    console.log('[PASS] Class 名に contextual keyword');
}

// ── AddressOf <contextual-keyword procedure> ──────────────────────────────────

{
    parse(`Function Step() As Long: Step = 1: End Function\nSub S(): Dim p As Long: p = AddressOf Step: End Sub`);
    parse(`Function Error() As Long: Error = 1: End Function\nSub S(): Dim p As Long: p = AddressOf Error: End Sub`);
    console.log('[PASS] AddressOf <contextual keyword 関数名>');
}

// ── TypeOf x Is <contextual-keyword type> ────────────────────────────────────

{
    parse(`Sub S()\nDim x As Object\nIf TypeOf x Is Error Then\nEnd If\nEnd Sub`);
    parse(`Sub S()\nDim x As Object\nIf TypeOf x Is Text Then\nEnd If\nEnd Sub`);
    parse(`Sub S()\nDim x As Object\nIf TypeOf x Is Step Then\nEnd If\nEnd Sub`);
    console.log('[PASS] TypeOf x Is <contextual keyword>');
}

// ── Implements <contextual-keyword> ──────────────────────────────────────────

{
    parse(`Implements Step`);
    parse(`Implements Error`);
    parse(`Implements Text`);
    console.log('[PASS] Implements <contextual keyword>');
}

// ── Class body: Public/Private <contextual-keyword> As Type ───────────────────
// §3.3.5.2: IDENTIFIER は contextual keyword を含む。
// `Public Step As Long` のようなフィールド宣言（Dim なし）が
// 無音でドロップされずに正しくパースされることを確認する。

{
    const { Lexer } = await import('../../src/engine/lexer');
    const { Parser } = await import('../../src/engine/parser');

    function getFields(code: string): string[] {
        const tokens = new Lexer(code).tokenize();
        const ast = new Parser(tokens, { errorRecovery: true }).parse();
        const cls = ast.body.find((s: any) => s.type === 'ClassDeclaration') as any;
        if (!cls) return [];
        return cls.body
            .filter((s: any) => s.type === 'VariableDeclaration')
            .flatMap((s: any) => s.declarations.map((d: any) => d.name.name));
    }

    assert.deepStrictEqual(getFields('Class C\nPublic Step As Long\nEnd Class'), ['Step'], 'Public Step フィールド');
    assert.deepStrictEqual(getFields('Class C\nPrivate Error As String\nEnd Class'), ['Error'], 'Private Error フィールド');
    assert.deepStrictEqual(getFields('Class C\nPublic Text As String\nEnd Class'), ['Text'], 'Public Text フィールド');
    console.log('[PASS] Class body: Public/Private <contextual keyword> As Type');
}

console.log('\n✅ contextual-kw-labels-goto: 全テスト通過');
