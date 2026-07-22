import { evalVBASingle, assert } from '../../test-libs/test-runner';
import { Lexer } from '../../src/engine/lexer';
import { Parser } from '../../src/engine/parser';

const ev = evalVBASingle(`
    Dim Shared moduleCounter As Long

    Function CountCall() As Long
        moduleCounter = moduleCounter + 1
        CountCall = moduleCounter
    End Function
`);

assert.strictEqual(ev.callProcedure('CountCall', []), 1);
assert.strictEqual(ev.callProcedure('CountCall', []), 2);

const ast = new Parser(new Lexer('Dim Shared value As Long').tokenize()).parse();
assert.strictEqual((ast.body[0] as any).isShared, true);

console.log('✅ Module-level Dim Shared is accepted with normal VBA semantics');
