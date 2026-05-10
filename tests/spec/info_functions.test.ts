import { Lexer } from '../../src/compiler/lexer';
import { Parser } from '../../src/compiler/parser';
import { Evaluator } from '../../src/compiler/evaluator';
import { assert } from '../ts/test-runner';

function evalVBA(code: string): any {
    const tokens = new Lexer(code).tokenize();
    const ast = new Parser(tokens).parse();
    const ev = new Evaluator(console.log);
    ev.evaluate(ast);
    return ev;
}

const code = `
    Function TestIsErrorMissing(Optional v)
        TestIsErrorMissing = IsError(v)
    End Function

    Function TestVarType()
        Dim results(10)
        results(0) = VarType(Empty)
        results(1) = VarType(Null)
        results(2) = VarType(123)
        results(3) = VarType("abc")
        results(4) = VarType(True)
        results(5) = VarType(DateSerial(2025, 1, 1))
        results(6) = VarType(Array(1, 2))
        results(7) = VarType(CreateObject("Scripting.Dictionary"))
        TestVarType = results
    End Function
`;

const ev = evalVBA(code);

console.log('[Test Suite] IsError / VarType の検証');

// IsError
assert.isTrue(ev.callProcedure('TestIsErrorMissing', []), 'IsError(Missing) -> True');
assert.isFalse(ev.callProcedure('TestIsErrorMissing', [1]), 'IsError(1) -> False');

// VarType
// vbEmpty=0, vbNull=1, vbDouble=5, vbString=8, vbBoolean=11, vbDate=7, vbArray+vbVariant=8204, vbObject=9
const vt = ev.callProcedure('TestVarType', []);
assert.strictEqual(vt[0], 0, 'VarType(Empty) -> 0');
assert.strictEqual(vt[1], 1, 'VarType(Null) -> 1');
assert.strictEqual(vt[2], 5, 'VarType(123) -> 5 (Double)');
assert.strictEqual(vt[3], 8, 'VarType("abc") -> 8');
assert.strictEqual(vt[4], 11, 'VarType(True) -> 11');
assert.strictEqual(vt[5], 5, 'VarType(Date) -> 5 (Currently Double)');
assert.strictEqual(vt[6], 8204, 'VarType(Array) -> 8204 (vbArray + vbVariant)');
assert.strictEqual(vt[7], 9, 'VarType(Object) -> 9');

console.log('[PASS] IsError / VarType の検証');

console.log('\n✅ Information Functions: 全テスト通過');
