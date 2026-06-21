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
