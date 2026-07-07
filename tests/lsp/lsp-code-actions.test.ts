import { LSPServer } from '../../src/lsp/server';
import { assert } from '../../test-libs/test-runner';

const URI = 'file:///test.bas';

function setup(code: string): LSPServer {
    const srv = new LSPServer();
    srv.didOpen(URI, code);
    return srv;
}

// 1. 選択範囲がプロシージャ外 → コードアクションなし
{
    const code = `Dim x As Long

Sub Foo()
End Sub`;
    const srv = setup(code);
    const actions = srv.getCodeActions(URI, {
        start: { line: 0, character: 0 },
        end:   { line: 0, character: 10 },
    });
    assert.strictEqual(actions.length, 0, '選択がプロシージャ外のときはアクションなし');
    console.log('[PASS] 選択範囲がプロシージャ外');
}

// 2. 変数ゼロの選択 → コードアクションなし
{
    const code = `Sub Foo()
    MsgBox "hello"
End Sub`;
    const srv = setup(code);
    const actions = srv.getCodeActions(URI, {
        start: { line: 1, character: 0 },
        end:   { line: 1, character: 0 },
    });
    assert.strictEqual(actions.length, 0, '変数ゼロはアクションなし');
    console.log('[PASS] 変数ゼロの選択');
}

// 3. inputs と outputs の分類
{
    // y は選択後でも使われるので outputs（ByRef）、x は選択前定義なので inputs（ByVal）
    const code = `Sub Foo()
    Dim x As Long
    x = 1
    Dim y As Long
    y = x + 1
    MsgBox y
End Sub`;
    // 選択: line 4 (y = x + 1) のみ (0-based)
    const srv = setup(code);
    const actions = srv.getCodeActions(URI, {
        start: { line: 4, character: 0 },
        end:   { line: 4, character: 20 },
    });
    assert.ok(actions.length > 0, 'コードアクションが返る');
    assert.ok(actions[0].command.command === 'vba-runner.doExtractFunction', 'コマンド名');
    const args = actions[0].command.arguments;
    assert.ok(args[2].inputs.includes('x'), 'x は ByVal 引数候補');
    assert.ok(args[2].outputs.includes('y'), 'y は ByRef 引数候補（選択後で使用）');
    console.log('[PASS] inputs と outputs の分類');
}

// 4. コードアクションのタイトルに呼び出し文が含まれる
{
    const code = `Sub Bar()
    Dim a As Long
    a = 1
    Dim b As Long
    b = a * 2
End Sub`;
    const srv = setup(code);
    const actions = srv.getCodeActions(URI, {
        start: { line: 4, character: 0 },
        end:   { line: 4, character: 10 },
    });
    assert.ok(actions.length > 0, 'アクションあり');
    assert.ok(actions[0].title.includes('Extract Function'), 'タイトルに Extract Function');
    console.log('[PASS] タイトルに呼び出し文');
}

// 5. outputs（ByRef 引数候補）と locals の分類
{
    const code = `Sub Calc()
    Dim total As Long
    Dim i As Integer
    Dim tmp As Long
    total = 0
    For i = 1 To 10
        tmp = i * 2
        total = total + tmp
    Next i
    MsgBox total
End Sub`;
    // 選択: line 4-8 (0-based) = total=0 〜 Next i
    const srv = setup(code);
    const actions = srv.getCodeActions(URI, {
        start: { line: 4, character: 0 },
        end:   { line: 8, character: 0 },
    });
    assert.ok(actions.length > 0, 'アクションあり');
    const result = actions[0].command.arguments[2];
    assert.ok(result.outputs.includes('total'), 'total は ByRef 引数候補（出力）');
    assert.ok(result.locals.includes('i') || result.locals.includes('tmp'), 'i か tmp はローカル変数候補');
    console.log('[PASS] outputs と locals の分類');
}

// 6. procSignature に正しいパラメーターが含まれる
{
    // result は選択後で使われるので ByRef 出力パラメーターになる
    const code = `Sub Process()
    Dim val As Long
    val = 42
    Dim result As Long
    result = val + 1
    MsgBox result
End Sub`;
    const srv = setup(code);
    const actions = srv.getCodeActions(URI, {
        start: { line: 4, character: 0 },
        end:   { line: 4, character: 20 },
    });
    assert.ok(actions.length > 0, 'アクションあり');
    const sig: string = actions[0].command.arguments[3];
    assert.ok(sig.includes('ByVal val'), 'procSignature に ByVal val');
    assert.ok(sig.includes('ByRef result'), 'procSignature に ByRef result');
    assert.ok(sig.startsWith('Private Sub'), 'Private Sub で始まる');
    assert.ok(sig.includes('End Sub'), 'End Sub で終わる');
    console.log('[PASS] procSignature のパラメーター');
}

// 7. callStatement の形式
{
    const code = `Sub Process()
    Dim val As Long
    val = 42
    Dim result As Long
    result = val + 1
    MsgBox result
End Sub`;
    const srv = setup(code);
    const actions = srv.getCodeActions(URI, {
        start: { line: 4, character: 0 },
        end:   { line: 4, character: 20 },
    });
    assert.ok(actions.length > 0, 'アクションあり');
    const call: string = actions[0].command.arguments[4];
    assert.ok(call.startsWith('ExtractedSub('), 'callStatement の形式');
    assert.ok(call.includes('val'), 'callStatement に val');
    assert.ok(call.includes('result'), 'callStatement に result');
    console.log('[PASS] callStatement の形式');
}

// 8. arguments の構成（uri, range, result, procSignature, callStatement）
{
    const code = `Sub Foo()
    Dim x As Long
    x = 1
    Dim y As Long
    y = x + 1
End Sub`;
    const srv = setup(code);
    const range = { start: { line: 4, character: 0 }, end: { line: 4, character: 10 } };
    const actions = srv.getCodeActions(URI, range);
    assert.ok(actions.length > 0, 'アクションあり');
    const args = actions[0].command.arguments;
    assert.strictEqual(args[0], URI, 'args[0] は URI');
    assert.deepStrictEqual(args[1], range, 'args[1] は range');
    assert.ok(typeof args[2] === 'object' && 'inputs' in args[2], 'args[2] は DefUseResult');
    assert.ok(typeof args[3] === 'string', 'args[3] は procSignature 文字列');
    assert.ok(typeof args[4] === 'string', 'args[4] は callStatement 文字列');
    console.log('[PASS] arguments の構成');
}

// 9. 行ドラッグ選択（end が次行の character 0）で End Sub 行を取り込まない
{
    const code = `Sub ResetBoard()
    Dim x As Integer
    Dim y As Integer
    For x = 0 To BOARD_WIDTH - 1
        For y = 0 To BOARD_HEIGHT - 1
            board(x, y) = 0
        Next y
    Next x
End Sub`;
    const srv = setup(code);
    // VS Code の行ドラッグ選択では「Next x」行末までのつもりでも
    // end が次の行（End Sub）の character 0 になることが多い
    const actions = srv.getCodeActions(URI, {
        start: { line: 3, character: 0 },
        end:   { line: 8, character: 0 },
    });
    assert.ok(actions.length > 0, 'アクションあり');
    const args = actions[0].command.arguments;
    const normalizedRange = args[1];
    assert.strictEqual(normalizedRange.end.line, 7, 'end.line は Next x の行（End Sub の手前）に正規化される');
    const sig: string = args[3];
    assert.ok(!sig.includes('End Sub\nEnd Sub'), 'procSignature に End Sub が二重出現しない');
    console.log('[PASS] 行ドラッグ選択で End Sub 行を取り込まない');
}

// 10. 正確な選択（オーバーシュートなし）では正規化が発生しない
{
    // Test 9 と同じコードを、行末まで character を正確に指定して選択（drag-select ではない通常の選択）
    const code = `Sub ResetBoard()
    Dim x As Integer
    Dim y As Integer
    For x = 0 To BOARD_WIDTH - 1
        For y = 0 To BOARD_HEIGHT - 1
            board(x, y) = 0
        Next y
    Next x
End Sub`;
    const srv = setup(code);
    const range = { start: { line: 3, character: 0 }, end: { line: 7, character: '    Next x'.length } };
    const actions = srv.getCodeActions(URI, range);
    assert.ok(actions.length > 0, 'アクションあり');
    const normalizedRange = actions[0].command.arguments[1];
    assert.deepStrictEqual(normalizedRange, range, 'character が0でない正確な選択は変更されない');
    console.log('[PASS] 正確な選択では正規化が発生しない');
}

// 11. Function でも行ドラッグ選択の End 行混入が防がれる（Sub 限定の修正ではないことの確認）
{
    const code = `Function SumBoard() As Long
    Dim x As Integer
    Dim total As Long
    For x = 0 To BOARD_WIDTH - 1
        total = total + board(x, 0)
    Next x
    SumBoard = total
End Function`;
    const srv = setup(code);
    // For ループ4行（For ... Next x）をドラッグ選択し、end が「End Function」直前の
    // 「SumBoard = total」行の次（character 0）になったケースを想定
    const actions = srv.getCodeActions(URI, {
        start: { line: 2, character: 0 },
        end:   { line: 6, character: 0 },
    });
    assert.ok(actions.length > 0, 'アクションあり');
    const args = actions[0].command.arguments;
    assert.strictEqual(args[1].end.line, 5, 'end.line は Next x の行に正規化される');
    const sig: string = args[3];
    assert.ok(!sig.includes('SumBoard = total'), 'procSignature に End Function 直前の文が混入しない');
    console.log('[PASS] Function でも行ドラッグ選択が正しく正規化される');
}

// buildExtractFunctionEdit が返す edit を実際に適用するヘルパー
// （vscode の WorkspaceEdit.replace + .insert と同じ意味で文字列を組み立てる）
function applyEditResult(code: string, editResult: {
    replaceRange: { startLine: number; endLine: number; endCharacter: number };
    replaceText: string;
    insertLine: number;
    insertText: string;
}): string {
    const lines = code.split('\n');
    const { replaceRange, replaceText, insertLine, insertText } = editResult;
    const before = lines.slice(0, replaceRange.startLine);
    const after  = lines.slice(replaceRange.endLine + 1);
    const offsetInAfter = insertLine - (replaceRange.endLine + 1);
    const afterHead = after.slice(0, offsetInAfter);
    const afterTail = after.slice(offsetInAfter);
    return [...before, replaceText, ...afterHead, ...insertText.split('\n'), ...afterTail].join('\n');
}

// 12. buildExtractFunctionEdit: 行ドラッグ選択（Test 9 と同じ正規化ケース）でも
//     For/Next の対応が壊れず、インデント・呼び出し文・End Sub の位置が正しい
{
    const code = `Sub ResetBoard()
    Dim x As Integer
    Dim y As Integer
    For x = 0 To BOARD_WIDTH - 1
        For y = 0 To BOARD_HEIGHT - 1
            board(x, y) = 0
        Next y
    Next x
End Sub`;
    const srv = setup(code);
    const actions = srv.getCodeActions(URI, {
        start: { line: 3, character: 0 },
        end:   { line: 8, character: 0 },
    });
    assert.ok(actions.length > 0, 'アクションあり');
    const [, range, result, procSignature, callStatement] = actions[0].command.arguments;
    const editResult = srv.buildExtractFunctionEdit(URI, range, 'ExtractedSub', result, procSignature, callStatement);
    assert.ok(editResult !== null, 'editResult が得られる');
    const finalText = applyEditResult(code, editResult!);
    assert.strictEqual(
        finalText,
        `Sub ResetBoard()
    Dim x As Integer
    Dim y As Integer
    ExtractedSub()
End Sub

Private Sub ExtractedSub()
    Dim x As Variant
    Dim y As Variant
    For x = 0 To BOARD_WIDTH - 1
        For y = 0 To BOARD_HEIGHT - 1
            board(x, y) = 0
        Next y
    Next x
End Sub
`,
        '抽出後のフォーマットが正しい（For/Next 対応・インデント・呼び出し文）',
    );
    console.log('[PASS] buildExtractFunctionEdit: 行ドラッグ選択でフォーマットが正しい');
}

// 13. buildExtractFunctionEdit: 選択範囲に含まれる Dim 宣言がパラメーターと衝突しない
//     Bug R1 regression: 選択行に "Dim x" が含まれ、かつ x が ByRef 出力パラメーターになる場合、
//     生成コードに "Dim x" と "ByRef x As Variant" が共存して VBA コンパイルエラーになっていた
{
    const code = `Sub ProcessData()
    Dim result As Long
    Dim x As Long
    x = 0
    x = x + 100
    result = x * 2
    Debug.Print result
End Sub`;
    const srv = setup(code);
    // "Dim x", "x = 0", "x = x + 100" を選択（0-based lines 2-4）
    // x は選択後の "result = x * 2" で使われるため ByRef 出力パラメーターになる
    const actions = srv.getCodeActions(URI, {
        start: { line: 2, character: 0 },
        end:   { line: 4, character: 20 },
    });
    assert.ok(actions.length > 0, 'アクションあり');
    const [, range, result2, procSignature, callStatement] = actions[0].command.arguments;
    assert.ok(result2.outputs.includes('x'), 'x は ByRef 出力候補');
    const editResult = srv.buildExtractFunctionEdit(URI, range, 'ProcessX', result2, procSignature, callStatement);
    assert.ok(editResult !== null, 'editResult が得られる');
    // 生成された新プロシージャに "Dim x" が重複しないこと（ByRef x はパラメーターで定義済み）
    assert.ok(!editResult!.insertText.toLowerCase().includes('dim x'), 'ByRef パラメーターの Dim 行が重複しない');
    assert.ok(editResult!.insertText.includes('ByRef x As Variant'), 'ByRef x は正しく引数として出力される');
    console.log('[PASS] buildExtractFunctionEdit: 選択内の Dim がパラメーターと衝突しない（Bug R1）');
}

// 14. buildExtractFunctionEdit: 選択範囲外で定義済みの変数には Dim を重複挿入しない、
//     かつ ByRef 出力（total）が引数として正しく扱われる
{
    const code = `Sub Calc()
    Dim total As Long
    Dim i As Integer
    Dim tmp As Long
    total = 0
    For i = 1 To 10
        tmp = i * 2
        total = total + tmp
    Next i
    MsgBox total
End Sub`;
    const srv = setup(code);
    // For ブロック全体（total=0 〜 Next i）を選択
    const actions = srv.getCodeActions(URI, {
        start: { line: 4, character: 0 },
        end:   { line: 9, character: 0 },
    });
    assert.ok(actions.length > 0, 'アクションあり');
    const [, range, result, procSignature, callStatement] = actions[0].command.arguments;
    const editResult = srv.buildExtractFunctionEdit(URI, range, 'ExtractedSub', result, procSignature, callStatement);
    assert.ok(editResult !== null, 'editResult が得られる');
    const finalText = applyEditResult(code, editResult!);
    assert.strictEqual(
        finalText,
        `Sub Calc()
    Dim total As Long
    Dim i As Integer
    Dim tmp As Long
    ExtractedSub(total)
    MsgBox total
End Sub

Private Sub ExtractedSub(ByRef total As Variant)
    Dim i As Variant
    Dim tmp As Variant
    total = 0
    For i = 1 To 10
        tmp = i * 2
        total = total + tmp
    Next i
End Sub
`,
        '抽出後のフォーマットが正しい（ByRef 引数・ローカル変数の Dim 挿入）',
    );
    console.log('[PASS] buildExtractFunctionEdit: ByRef 引数とローカル変数の Dim 挿入が正しい');
}
