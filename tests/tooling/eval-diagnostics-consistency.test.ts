import { spawnSync } from 'node:child_process';
import { assert, evalVBASingle, evalVBAModules } from '../../test-libs/test-runner';

// The default keeps valid engine tests aligned with the published LSP path.
{
    const source = String.raw`Option Explicit
Function AddOne(ByVal value As Long) As Long
    AddOne = value + 1
End Function`;
    const evaluator = evalVBASingle(source);
    assert.strictEqual(evaluator.callProcedure('AddOne', [4]), 5,
        'valid source remains executable with the default consistency check');
    console.log('[PASS] evalVBASingle default diagnostics consistency');
}

// An explicit LSP-error expectation is independent of engine call timing.
{
    const source = String.raw`Option Explicit
Sub BadSub()
    missing = 1
End Sub`;
    const evaluator = evalVBASingle(source);
    assert.throws(() => evaluator.callProcedure('BadSub', []), /not declared/,
        'the caught exception remains the engine error');
    console.log('[PASS] LSP error expectation preserves engine exception flow');
}

// A mismatch is deferred and changes the process status after the test body,
// rather than being thrown from evalVBASingle.
{
    const child = spawnSync(process.execPath, [
        '--import', 'tsx', '--input-type=module', '-e',
        "import { evalVBASingle } from './test-libs/test-runner.ts'; const ev = evalVBASingle(String.raw`Option Explicit\nFunction GoodFunc() As Long\n    GoodFunc = 99\nEnd Function`, { diagnostics: { expectation: 'lsp-error' } }); ev.callProcedure('GoodFunc', []);",
    ], { cwd: process.cwd(), encoding: 'utf8' });
    assert.notStrictEqual(child.status, 0,
        'a deferred LSP consistency mismatch must fail the process');
    console.log('[PASS] Deferred LSP mismatch fails the test process');
}

// GoodFunc remains callable while the same module explicitly expects the
// static LSP error in BadSub.
{
    const modules = [{ name: 'Module', code: String.raw`Option Explicit
Function GoodFunc() As Long
    GoodFunc = 99
End Function
Sub BadSub()
    missing = 1
End Sub` }];
    const evaluator = evalVBAModules(modules);
    assert.strictEqual(evaluator.callProcedure('GoodFunc', []), 99,
        'a valid procedure remains executable when another procedure has an LSP error');
    assert.throws(() => evaluator.callProcedure('BadSub', []), /not declared/,
        'the bad procedure still raises the engine error when called');
    console.log('[PASS] GoodFunc/BadSub LSP expectation matches engine timing');
}

// A preproc error raised by Inner is rethrown while Outer unwinds.  The
// consistency check must attribute it to Inner only; checking Outer would
// produce a false "has no LSP Error" mismatch.
{
    const modules = [{ name: 'Module', code: String.raw`Option Explicit
Sub Outer()
    Inner
End Sub
Private Sub Inner()
    missing = 1
End Sub` }];
    const evaluator = evalVBAModules(modules);
    assert.throws(() => evaluator.callProcedure('Outer', []), /not declared/,
        'a nested preproc error remains the engine exception');
    console.log('[PASS] Nested preproc errors are attributed to the originating procedure');
}

// The same attribution rule applies across modules, where the module name is
// needed in addition to the procedure name to select the correct diagnostic.
{
    const modules = [
        { name: 'Caller', code: String.raw`Sub Start()
    Worker.DoWork
End Sub` },
        { name: 'Worker', code: String.raw`Option Explicit
Public Sub DoWork()
    missing = 1
End Sub` },
    ];
    const evaluator = evalVBAModules(modules);
    assert.throws(() => evaluator.callProcedure('Start', []), /not declared/,
        'a cross-module nested preproc error remains the engine exception');
    console.log('[PASS] Cross-module nested preproc errors keep their origin');
}

// Exceptional harness-only sources may opt out, but the reason is mandatory.
{
    const source = String.raw`Sub Verify()
End Sub`;
    const evaluator = evalVBASingle(source, {
        diagnostics: { expectation: 'skip', reason: 'exercise the explicit opt-out contract' },
    });
    evaluator.callProcedure('Verify', []);
    assert.throws(() => evalVBASingle(source, {
        diagnostics: { expectation: 'skip', reason: '' },
    }), /non-empty reason/, 'diagnostic opt-out requires a reason');
    console.log('[PASS] Diagnostic opt-out requires an explicit reason');
}

console.log('\n✅ Evaluation diagnostics consistency: all tests passed');
