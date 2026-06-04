/**
 * VBADebugSession 統合テスト（Worker スレッド + Atomics）
 *
 * 実際に Worker を起動して DAP ステップ実行の動作を検証する。
 * tsx 実行時は debug-worker.ts を tsx ローダーで直接起動する。
 */
import { VBADebugSession } from '../../src/lsp/debug-session';
import { assert } from '../../test-libs/test-runner';

const TIMEOUT_MS = 8000;

/** session から指定イベントが来るまで待つ */
function waitFor(session: VBADebugSession, event: string): Promise<any> {
    return new Promise((resolve, reject) => {
        const timer = setTimeout(
            () => reject(new Error(`Timeout (${TIMEOUT_MS}ms) waiting for '${event}'`)),
            TIMEOUT_MS
        );
        session.once(event, (data: any) => {
            clearTimeout(timer);
            resolve(data);
        });
    });
}

/** session を開始して最初の stopped を待つ */
async function startAndWait(session: VBADebugSession): Promise<{ line: number; reason: string }> {
    const p = waitFor(session, 'stopped');
    session.start();
    return p;
}

// ──────────────────────────────────────────────
// 1. セッション起動で最初のステートメントに停止する
// ──────────────────────────────────────────────
{
    const src = [
        'Sub Main()',
        '  Dim x As Long',
        '  x = 1',
        '  x = 2',
        'End Sub',
    ].join('\n');

    const session = new VBADebugSession(src, 'Module1');
    const info = await startAndWait(session);

    assert.ok(typeof info.line === 'number', 'stopped event has line');
    assert.strictEqual(info.reason, 'entry', 'first stop reason is entry');
    assert.ok(info.line >= 2, 'paused at a statement line (>= 2)');

    session.terminate();
    await waitFor(session, 'exited').catch(() => { /* ok */ });
    console.log('[PASS] Session starts and pauses at entry, line:', info.line);
}

// ──────────────────────────────────────────────
// 2. stepInto を繰り返すと行が進む
// ──────────────────────────────────────────────
{
    const src = [
        'Sub Main()',
        '  Dim x As Long',
        '  x = 1',
        '  x = 2',
        '  x = 3',
        'End Sub',
    ].join('\n');

    const session = new VBADebugSession(src, 'Module1');
    const first = await startAndWait(session);

    const second = await ((): Promise<any> => {
        const p = waitFor(session, 'stopped');
        session.stepInto();
        return p;
    })();

    assert.ok(second.line > first.line || second.line === first.line,
        'stepped to next or same line');
    assert.strictEqual(second.reason, 'step', 'reason is step after stepInto');

    const third = await ((): Promise<any> => {
        const p = waitFor(session, 'stopped');
        session.stepInto();
        return p;
    })();

    assert.ok(third.line >= second.line, 'third step line >= second');

    session.terminate();
    await waitFor(session, 'exited').catch(() => { /* ok */ });
    console.log('[PASS] stepInto advances through statements:', first.line, '→', second.line, '→', third.line);
}

// ──────────────────────────────────────────────
// 3. stepOver はプロシージャ呼び出しを飛び越える
// ──────────────────────────────────────────────
{
    const src = [
        'Sub Main()',      // line 1
        '  Dim x As Long', // line 2
        '  x = 1',         // line 3
        '  Call Helper()', // line 4
        '  x = 2',         // line 5
        'End Sub',         // line 6
        'Sub Helper()',    // line 7
        '  Dim y As Long', // line 8
        '  y = 999',       // line 9
        'End Sub',         // line 10
    ].join('\n');

    const session = new VBADebugSession(src, 'Module1');
    await startAndWait(session); // entry pause

    // 最初の停止位置から Call Helper() まで stepInto で進む
    let current = session.getCurrentLine();
    let maxSteps = 10;
    while (current !== 4 && maxSteps-- > 0) {
        const p = waitFor(session, 'stopped');
        session.stepInto();
        const info = await p;
        current = info.line;
    }
    assert.strictEqual(current, 4, 'reached Call Helper() at line 4');

    // Call Helper() で stepOver → Helper の中には入らず x = 2 (line 5) に停止
    const afterOver = await ((): Promise<any> => {
        const p = waitFor(session, 'stopped');
        session.stepOver();
        return p;
    })();

    assert.ok(afterOver.line !== 8 && afterOver.line !== 9,
        `stepOver did not enter Helper (line ${afterOver.line})`);
    assert.ok(afterOver.line >= 5, 'stepOver landed after Call Helper()');

    session.terminate();
    await waitFor(session, 'exited').catch(() => { /* ok */ });
    console.log('[PASS] stepOver skips Helper body, landed at line:', afterOver.line);
}

// ──────────────────────────────────────────────
// 4. stepInto はプロシージャ呼び出しに入る
// ──────────────────────────────────────────────
{
    const src = [
        'Sub Main()',      // line 1
        '  Call Inner()', // line 2
        'End Sub',         // line 3
        'Sub Inner()',     // line 4
        '  Dim z As Long', // line 5
        '  z = 42',        // line 6
        'End Sub',         // line 7
    ].join('\n');

    const session = new VBADebugSession(src, 'Module1');
    await startAndWait(session); // entry: line 2 (Call Inner())

    // Call Inner() で stepInto → Inner の中に入る
    const inInner = await ((): Promise<any> => {
        const p = waitFor(session, 'stopped');
        session.stepInto();
        return p;
    })();

    // Inner の中のステートメント（line 5 or 6）に停止しているはず
    assert.ok(inInner.line >= 5 && inInner.line <= 6,
        `stepInto entered Inner at line ${inInner.line}`);

    session.terminate();
    await waitFor(session, 'exited').catch(() => { /* ok */ });
    console.log('[PASS] stepInto enters Inner procedure at line:', inInner.line);
}

// ──────────────────────────────────────────────
// 5. stepOut は現在のプロシージャを抜ける
// ──────────────────────────────────────────────
{
    const src = [
        'Sub Main()',      // line 1
        '  Call Inner()', // line 2
        '  Dim a As Long', // line 3
        '  a = 0',         // line 4
        'End Sub',         // line 5
        'Sub Inner()',     // line 6
        '  Dim z As Long', // line 7
        '  z = 1',         // line 8
        '  z = 2',         // line 9
        'End Sub',         // line 10
    ].join('\n');

    const session = new VBADebugSession(src, 'Module1');
    await startAndWait(session);

    // Call Inner() に到達するまで stepInto
    let cur = session.getCurrentLine();
    let steps = 10;
    while (cur !== 2 && steps-- > 0) {
        const p = waitFor(session, 'stopped');
        session.stepInto();
        const i = await p;
        cur = i.line;
    }
    assert.strictEqual(cur, 2, 'reached Call Inner() at line 2');

    // Inner に入る
    const inInner = await ((): Promise<any> => {
        const p = waitFor(session, 'stopped');
        session.stepInto();
        return p;
    })();
    assert.ok(inInner.line >= 7, `entered Inner at line ${inInner.line}`);

    // Inner から stepOut → Main に戻る
    const backInMain = await ((): Promise<any> => {
        const p = waitFor(session, 'stopped');
        session.stepOut();
        return p;
    })();
    assert.ok(backInMain.line >= 3 && backInMain.line <= 4,
        `stepOut returned to Main at line ${backInMain.line}`);

    session.terminate();
    await waitFor(session, 'exited').catch(() => { /* ok */ });
    console.log('[PASS] stepOut returns to Main at line:', backInMain.line);
}

// ──────────────────────────────────────────────
// 6. continue でプログラムを最後まで実行する
// ──────────────────────────────────────────────
{
    const src = [
        'Sub Main()',
        '  Dim x As Long',
        '  x = 1',
        '  x = 2',
        'End Sub',
    ].join('\n');

    const session = new VBADebugSession(src, 'Module1');
    await startAndWait(session);

    const exitedP = waitFor(session, 'exited');
    session.continue();
    const exitCode = await exitedP;

    assert.strictEqual(exitCode, 0, 'program exited with code 0');
    console.log('[PASS] continue runs program to completion');
}

// ──────────────────────────────────────────────
// 7. ブレークポイント行で停止する
// ──────────────────────────────────────────────
{
    const src = [
        'Sub Main()',
        '  Dim x As Long',
        '  x = 1',         // line 3
        '  x = 2',         // line 4  ← breakpoint
        '  x = 3',         // line 5
        'End Sub',
    ].join('\n');

    const session = new VBADebugSession(src, 'Module1');
    session.setBreakpoints([4]); // line 4 にブレークポイント
    await startAndWait(session); // entry pause

    // continue で走らせてブレークポイントに停止するのを待つ
    const atBp = await ((): Promise<any> => {
        const p = waitFor(session, 'stopped');
        session.continue();
        return p;
    })();

    assert.strictEqual(atBp.line, 4, 'stopped at breakpoint line 4');
    assert.strictEqual(atBp.reason, 'breakpoint', 'reason is breakpoint');

    session.terminate();
    await waitFor(session, 'exited').catch(() => { /* ok */ });
    console.log('[PASS] Breakpoint stops execution at line 4');
}

// ──────────────────────────────────────────────
// 8. 停止時に変数値が正しく取れる
// ──────────────────────────────────────────────
{
    const src = [
        'Sub Main()',
        '  Dim x As Long',
        '  x = 123',        // line 3
        '  x = 456',        // line 4  ← breakpoint: x=123 が見える
        'End Sub',
    ].join('\n');

    const session = new VBADebugSession(src, 'Module1');
    session.setBreakpoints([4]);
    await startAndWait(session);

    const atBp = await ((): Promise<any> => {
        const p = waitFor(session, 'stopped');
        session.continue();
        return p;
    })();

    assert.strictEqual(atBp.line, 4, 'stopped at line 4');

    const vars = session.getVariables(0);
    const xVar = vars.find(v => v.name === 'x');
    assert.ok(xVar !== undefined, 'variable x found');
    assert.strictEqual(xVar!.value, '123', 'x = 123 at breakpoint line 4');

    session.terminate();
    await waitFor(session, 'exited').catch(() => { /* ok */ });
    console.log('[PASS] Variables show correct values at breakpoint');
}

// ──────────────────────────────────────────────
// 9. スタックフレームが取れる
// ──────────────────────────────────────────────
{
    const src = [
        'Sub Main()',
        '  Call Foo()',
        'End Sub',
        'Sub Foo()',
        '  Dim y As Long',
        '  y = 1',
        'End Sub',
    ].join('\n');

    const session = new VBADebugSession(src, 'Module1');
    await startAndWait(session);

    // Foo に入る
    let cur = session.getCurrentLine();
    let steps = 5;
    while (cur !== 2 && steps-- > 0) {
        const p = waitFor(session, 'stopped');
        session.stepInto();
        const i = await p;
        cur = i.line;
    }
    await ((): Promise<any> => {
        const p = waitFor(session, 'stopped');
        session.stepInto();
        return p;
    })();

    const frames = session.getStackFrames();
    assert.ok(frames.length >= 1, 'has at least 1 frame');
    assert.ok(frames.some(f => f.name === 'Foo'), 'Foo frame present');

    session.terminate();
    await waitFor(session, 'exited').catch(() => { /* ok */ });
    console.log('[PASS] Stack frames available inside called procedure');
}

// ──────────────────────────────────────────────
// 10. 出力が output イベントで通知される
// ──────────────────────────────────────────────
{
    const src = [
        'Sub Main()',
        '  Debug.Print "hello"',
        'End Sub',
    ].join('\n');

    const session = new VBADebugSession(src, 'Module1');
    const outputs: string[] = [];
    session.on('output', (text: string) => outputs.push(text));
    await startAndWait(session);

    const exitedP = waitFor(session, 'exited');
    session.continue();
    await exitedP;

    assert.ok(outputs.some(o => o.includes('hello')), 'output contains "hello"');
    console.log('[PASS] Debug.Print output received via output event');
}

console.log('\n✅ Debug Session: 全テスト通過');
