import { DebugAdapter } from '../../src/lsp/debug-adapter';
import { assert } from '../../test-libs/test-runner';

function createAdapter(src: string): DebugAdapter {
    return new DebugAdapter(src, 'TestModule');
}

// 1. Initialize request
{
    const code = 'Sub Test()\nEnd Sub';
    const adapter = createAdapter(code);
    const response = adapter.handleInitialize();
    assert.ok(response, 'initialize response');
    assert.ok(response.capabilities, 'capabilities present');
    console.log('[PASS] Initialize request');
}

// 2. Launch request
{
    const code = 'Sub Test()\nEnd Sub';
    const adapter = createAdapter(code);
    adapter.handleInitialize();
    const response = adapter.handleLaunch({});
    assert.ok(response !== null, 'launch response');
    assert.ok(response.success, 'launch succeeded');
    // worker が起動するので後始末
    adapter.handleDisconnect();
    console.log('[PASS] Launch request');
}

// 3. Set breakpoints request
{
    const code = 'Sub Test()\n  x = 1\nEnd Sub';
    const adapter = createAdapter(code);
    adapter.handleInitialize();
    const response = adapter.handleSetBreakpoints({ breakpoints: [{ line: 2, column: 0, condition: 'x = 1' }] });
    assert.ok(response, 'breakpoints response');
    assert.ok(Array.isArray(response.breakpoints), 'breakpoints array');
    assert.strictEqual(response.breakpoints[0].condition, 'x = 1', 'condition is preserved');
    console.log('[PASS] Set breakpoints request');
}

// 4. Threads request
{
    const code = 'Sub Test()\nEnd Sub';
    const adapter = createAdapter(code);
    adapter.handleInitialize();
    const response = adapter.handleThreads();
    assert.ok(Array.isArray(response), 'threads array');
    assert.ok(response.length > 0, 'at least one thread');
    assert.strictEqual(response[0].id, 1, 'thread id is 1');
    console.log('[PASS] Threads request');
}

// 5. Stack trace request
{
    const code = 'Sub Test()\nEnd Sub';
    const adapter = createAdapter(code);
    adapter.handleInitialize();
    const response = adapter.handleStackTrace(1);
    assert.ok(response, 'stack trace response');
    assert.ok(Array.isArray(response.stackFrames), 'stack frames array');
    console.log('[PASS] Stack trace request');
}

// 6. Variables request
{
    const code = 'Sub Test()\n  Dim x As Integer\nEnd Sub';
    const adapter = createAdapter(code);
    adapter.handleInitialize();
    const response = adapter.handleVariables(0);
    assert.ok(response, 'variables response');
    assert.ok(Array.isArray(response.variables), 'variables array');
    console.log('[PASS] Variables request');
}

// 7. Continue request
{
    const code = 'Sub Test()\nEnd Sub';
    const adapter = createAdapter(code);
    adapter.handleInitialize();
    const response = adapter.handleContinue(1);
    assert.ok(response !== null, 'continue returns response');
    console.log('[PASS] Continue request');
}

// 8. Step over request
{
    const code = 'Sub Test()\nEnd Sub';
    const adapter = createAdapter(code);
    adapter.handleInitialize();
    const response = adapter.handleStepOver(1);
    assert.ok(response !== null, 'step over response');
    console.log('[PASS] Step over request');
}

// 9. Evaluate request while paused
await (async () => {
    const code = 'Sub Test()\n  Dim x As Long\n  x = 41\nEnd Sub';
    const adapter = createAdapter(code);
    adapter.handleInitialize();
    const stopped = new Promise<void>((resolve) => {
        adapter.onEvent = (event) => {
            if (event.event === 'stopped') resolve();
        };
    });
    adapter.handleLaunch({});
    await stopped;
    const response = await adapter.handleEvaluate('1 + 1', 0);
    assert.ok(response, 'evaluate response');
    assert.strictEqual(response.result, '2', 'evaluated expression result');
    assert.strictEqual(response.type, 'Long', 'evaluated expression type');
    adapter.handleDisconnect();
    console.log('[PASS] Evaluate request while paused');
})();

// 10. setVariable request while paused
await (async () => {
    const code = 'Sub Test()\n  Dim x As Long\n  x = 1\n  Debug.Print x\nEnd Sub';
    const adapter = createAdapter(code);
    adapter.handleInitialize();
    const stopped = new Promise<void>((resolve) => {
        adapter.onEvent = (event) => {
            if (event.event === 'stopped') resolve();
        };
    });
    adapter.handleLaunch({});
    await stopped;
    adapter.handleSetBreakpoints({ breakpoints: [{ line: 4, column: 0 }] });
    // launch enters at line 2; continue to the configured breakpoint.
    const bpStopped = new Promise<void>((resolve) => {
        adapter.onEvent = (event) => {
            if (event.event === 'stopped') resolve();
        };
    });
    adapter.handleContinue(1);
    await bpStopped;
    const response = await adapter.handleSetVariable(0, 'x', '42');
    assert.strictEqual(response.result, '42', 'setVariable returns assigned value');
    assert.strictEqual(response.type, 'Long', 'setVariable preserves VBA type');
    adapter.handleDisconnect();
    console.log('[PASS] Set variable request while paused');
})();

// 11. Disconnect request
{
    const code = 'Sub Test()\nEnd Sub';
    const adapter = createAdapter(code);
    adapter.handleInitialize();
    adapter.handleLaunch({});
    const response = adapter.handleDisconnect();
    assert.ok(response !== null, 'disconnect response');
    console.log('[PASS] Disconnect request');
}

console.log('\n✅ LSP Debug Adapter: 全テスト通過');
