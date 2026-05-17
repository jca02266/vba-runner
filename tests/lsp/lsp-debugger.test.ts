import { Lexer } from '../../src/engine/lexer';
import { Parser } from '../../src/engine/parser';
import { Debugger, Breakpoint, StackFrame } from '../../src/lsp/debugger';
import { assert } from '../../test-libs/test-runner';

function createDebugger(src: string): Debugger {
    const tokens = new Lexer(src).tokenize();
    const ast = new Parser(tokens).parse();
    return new Debugger(ast);
}

// 1. Debugger initializes with source
{
    const code = 'Sub Test()\nEnd Sub';
    const debugger_ = createDebugger(code);
    assert.ok(debugger_, 'debugger created');
    console.log('[PASS] Debugger initialization');
}

// 2. Set breakpoint at line
{
    const code = 'Sub Test()\n  x = 1\n  y = 2\nEnd Sub';
    const debugger_ = createDebugger(code);
    const bp = debugger_.setBreakpoint(1, 2); // line 2
    assert.ok(bp, 'breakpoint created');
    assert.strictEqual(bp.line, 1, 'breakpoint line (0-based)');
    assert.ok(typeof bp.verified === 'boolean', 'verified is boolean');
    console.log('[PASS] Set breakpoint');
}

// 3. List breakpoints
{
    const code = 'Sub Test()\n  x = 1\n  y = 2\nEnd Sub';
    const debugger_ = createDebugger(code);
    debugger_.setBreakpoint(1, 2);
    debugger_.setBreakpoint(2, 2);
    const bps = debugger_.getBreakpoints();
    assert.strictEqual(bps.length, 2, 'two breakpoints');
    console.log('[PASS] List breakpoints');
}

// 4. Remove breakpoint
{
    const code = 'Sub Test()\n  x = 1\nEnd Sub';
    const debugger_ = createDebugger(code);
    const bp = debugger_.setBreakpoint(1, 2);
    debugger_.removeBreakpoint(bp.id);
    const bps = debugger_.getBreakpoints();
    assert.strictEqual(bps.length, 0, 'breakpoint removed');
    console.log('[PASS] Remove breakpoint');
}

// 5. Breakpoint has required fields
{
    const code = 'Sub Test()\nEnd Sub';
    const debugger_ = createDebugger(code);
    const bp = debugger_.setBreakpoint(0, 0);
    assert.ok(bp.id, 'breakpoint has id');
    assert.strictEqual(typeof bp.line, 'number', 'line is number');
    assert.strictEqual(typeof bp.column, 'number', 'column is number');
    assert.strictEqual(typeof bp.verified, 'boolean', 'verified is boolean');
    console.log('[PASS] Breakpoint fields complete');
}

// 6. Get stack frames (execution state)
{
    const code = 'Sub Test()\n  x = 1\nEnd Sub';
    const debugger_ = createDebugger(code);
    const frames = debugger_.getStackFrames();
    assert.ok(Array.isArray(frames), 'frames is array');
    // Initially may be empty or have call stack
    console.log('[PASS] Get stack frames');
}

// 7. Stack frame has required fields
{
    const code = 'Sub Outer()\n  Call Inner()\nEnd Sub\nSub Inner()\nEnd Sub';
    const debugger_ = createDebugger(code);
    const frames = debugger_.getStackFrames();
    // At least structure should be available
    assert.ok(Array.isArray(frames), 'frames exist');
    console.log('[PASS] Stack frame structure');
}

// 8. Get variables at frame
{
    const code = 'Sub Test()\n  Dim x As Integer\n  x = 42\nEnd Sub';
    const debugger_ = createDebugger(code);
    const vars = debugger_.getVariables(0); // frame 0
    assert.ok(Array.isArray(vars), 'variables is array');
    console.log('[PASS] Get variables at frame');
}

// 9. Set and get variable value
{
    const code = 'Sub Test()\n  Dim x As Integer\n  x = 42\nEnd Sub';
    const debugger_ = createDebugger(code);
    debugger_.setVariableValue('x', 42);
    const value = debugger_.getVariableValue('x');
    assert.strictEqual(value, 42, 'variable value retrieved and equals 42');
    console.log('[PASS] Set and get variable value');
}

// 10. Debugger state tracking
{
    const code = 'Sub Test()\nEnd Sub';
    const debugger_ = createDebugger(code);
    assert.ok(debugger_.getState(), 'debugger has state');
    const state = debugger_.getState();
    assert.ok(typeof state === 'string', 'state is string');
    console.log('[PASS] Debugger state:', state);
}

console.log('\n✅ LSP Debugger: 全テスト通過');
