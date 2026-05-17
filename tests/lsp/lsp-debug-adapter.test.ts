import { Lexer } from '../../src/engine/lexer';
import { Parser } from '../../src/engine/parser';
import { DebugAdapter } from '../../src/lsp/debug-adapter';
import { assert } from '../../test-libs/test-runner';

function createAdapter(src: string): DebugAdapter {
    const tokens = new Lexer(src).tokenize();
    const ast = new Parser(tokens).parse();
    return new DebugAdapter(ast);
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
    console.log('[PASS] Launch request');
}

// 3. Set breakpoints request
{
    const code = 'Sub Test()\n  x = 1\nEnd Sub';
    const adapter = createAdapter(code);
    adapter.handleInitialize();
    const response = adapter.handleSetBreakpoints({ line: 1, column: 0 });
    assert.ok(Array.isArray(response), 'breakpoints array');
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
    console.log('[PASS] Threads request');
}

// 5. Stack trace request
{
    const code = 'Sub Test()\nEnd Sub';
    const adapter = createAdapter(code);
    adapter.handleInitialize();
    const response = adapter.handleStackTrace(1); // threadId
    assert.ok(Array.isArray(response), 'stack frames array');
    console.log('[PASS] Stack trace request');
}

// 6. Variables request
{
    const code = 'Sub Test()\n  Dim x As Integer\nEnd Sub';
    const adapter = createAdapter(code);
    adapter.handleInitialize();
    const response = adapter.handleVariables(0); // frameId
    assert.ok(Array.isArray(response), 'variables array');
    console.log('[PASS] Variables request');
}

// 7. Continue request
{
    const code = 'Sub Test()\nEnd Sub';
    const adapter = createAdapter(code);
    adapter.handleInitialize();
    const response = adapter.handleContinue(1);
    assert.ok(typeof response === 'boolean', 'continue returns boolean');
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

// 9. Evaluate request
{
    const code = 'Sub Test()\nEnd Sub';
    const adapter = createAdapter(code);
    adapter.handleInitialize();
    const response = adapter.handleEvaluate('x + 1', 0); // expression, frameId
    assert.ok(response, 'evaluate response');
    assert.ok(typeof response.result === 'string', 'result is string');
    console.log('[PASS] Evaluate request');
}

// 10. Disconnect request
{
    const code = 'Sub Test()\nEnd Sub';
    const adapter = createAdapter(code);
    adapter.handleInitialize();
    const response = adapter.handleDisconnect();
    assert.ok(response !== null, 'disconnect response');
    console.log('[PASS] Disconnect request');
}

console.log('\n✅ LSP Debug Adapter: 全テスト通過');
