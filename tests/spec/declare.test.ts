import { assert } from '../ts/test-runner';
import { Evaluator } from '../../src/compiler/evaluator';
import { Parser } from '../../src/compiler/parser';
import { Lexer } from '../../src/compiler/lexer';

function evalVBA(code: string) {
    const tokens = new Lexer(code).tokenize();
    const parser = new Parser(tokens);
    const program = parser.parse();
    const evaluator = new Evaluator((s) => console.log(s));
    evaluator.evaluate(program);
    return evaluator;
}

console.log('[Test Suite] Declare Statement (External Function) の検証');

const code = `
    Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long
    Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
    
    Function TestDeclare()
        Dim t
        t = GetTickCount()
        Sleep 100
        TestDeclare = t
    End Function
`;

const ev = evalVBA(code);

// Should not throw and should call the stub
assert.strictEqual(ev.callProcedure('TestDeclare', []), 0, 'Declare function should return 0 (stub default)');

console.log('✅ Declare Statement: 全テスト通過');
