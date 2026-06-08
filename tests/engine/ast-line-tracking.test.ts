import { Lexer } from '../../src/engine/lexer';
import { Parser } from '../../src/engine/parser';
import { evalVBASingle, assert } from '../../test-libs/test-runner';

function runVba(code: string) {
    return evalVBASingle(code, { onPrint: () => {} });
}

function catchError(code: string, proc: string): any {
    const ev = runVba(code);
    try {
        ev.callProcedure(proc, []);
        return null;
    } catch (e: any) {
        return e;
    }
}

// 1. Parser attaches line numbers to statements
{
    const source = `
Sub Dummy()
    Dim x
    x = 1
End Sub
`;
    const tokens = new Lexer(source).tokenize();
    const ast = new Parser(tokens).parse();
    const procDecl = ast.body.find((s: any) => s.type === 'ProcedureDeclaration') as any;
    assert.ok(procDecl !== undefined, 'ProcedureDeclaration found');
    assert.ok(procDecl.line !== undefined, 'ProcedureDeclaration has line');
    assert.ok(procDecl.line > 0, 'line > 0');

    const assignStmt = procDecl.body.find((s: any) => s.type === 'AssignmentStatement') as any;
    assert.ok(assignStmt !== undefined, 'AssignmentStatement found');
    assert.ok(assignStmt.line !== undefined, 'AssignmentStatement has line');
    assert.ok(assignStmt.line > procDecl.line, 'AssignmentStatement line > proc declaration line');
    console.log('[PASS] Parser attaches line numbers; proc=', procDecl.line, 'assign=', assignStmt.line);
}

// 2. Type mismatch error includes line number in message
{
    const source = `
Sub Test2()
    Dim x As Integer
    x = CInt("abc")
End Sub
`;
    const err = catchError(source, 'Test2');
    assert.ok(err !== null, 'Should throw error');
    assert.ok(err.vbaLine !== undefined, 'Error should have vbaLine');
    assert.ok(err.message.includes('line'), 'Error message includes "line"');
    assert.ok(err.vbaLine > 0, 'vbaLine is positive');
    console.log('[PASS] Type mismatch vbaLine:', err.vbaLine, 'message:', err.message);
}

// 3. Line 6 error (multi-line procedure)
{
    const source = `
Sub Test3()
    Dim a As Integer
    Dim b As Integer
    Dim c As Integer
    c = CInt("bad value here")
End Sub
`;
    const err = catchError(source, 'Test3');
    assert.ok(err !== null, 'Should throw on line 6');
    assert.ok(err.vbaLine !== undefined, 'vbaLine is set');
    assert.strictEqual(err.vbaLine, 6, 'vbaLine is 6');
    console.log('[PASS] Multi-line proc error on line 6, vbaLine:', err.vbaLine);
}

// 4. Division by zero error includes line
{
    const source = `
Sub Test4()
    Dim x As Integer
    Dim y As Integer
    y = 0
    x = 10 / y
End Sub
`;
    const err = catchError(source, 'Test4');
    assert.ok(err !== null, 'Should throw division by zero');
    assert.ok(err.vbaLine !== undefined, 'vbaLine set for division by zero');
    assert.ok(err.vbaLine > 0, 'vbaLine > 0');
    console.log('[PASS] Division by zero vbaLine:', err.vbaLine);
}

// 5. Line numbers increase correctly across multiple statements
{
    const source = `
Sub Foo()
    Dim a
    Dim b
    Dim c
    a = 1
    b = 2
    c = 3
End Sub
`;
    const tokens = new Lexer(source).tokenize();
    const ast = new Parser(tokens).parse();
    const proc = ast.body.find((s: any) => s.type === 'ProcedureDeclaration') as any;
    const stmts = proc.body.filter((s: any) => s.line !== undefined);
    assert.ok(stmts.length > 0, 'Body statements have lines');
    for (let i = 1; i < stmts.length; i++) {
        assert.ok(stmts[i].line >= stmts[i-1].line, 'Lines are non-decreasing');
    }
    console.log('[PASS] Line numbers non-decreasing across statements');
}

// 6. vbaLine property is accessible on error object
{
    const source = `
Sub Test6()
    Dim x As Double
    x = CDbl("not-a-number")
End Sub
`;
    const err = catchError(source, 'Test6');
    assert.ok(err !== null, 'Should throw');
    assert.ok('vbaLine' in err, 'err has vbaLine property');
    assert.ok(err.type === 'VbaError', 'err.type is VbaError');
    console.log('[PASS] vbaLine property on error, type:', err.type, 'line:', err.vbaLine);
}

console.log('\n✅ AST line tracking: 全テスト通過');
