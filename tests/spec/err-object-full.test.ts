/**
 * Err Object (§6.1.3.2) のテスト
 */
import { Lexer } from '../../src/engine/lexer';
import { Parser } from '../../src/engine/parser';
import { Evaluator } from '../../src/engine/evaluator';
import { assert } from '../../test-libs/test-runner';

function evalVBA(code: string): any {
    const tokens = new Lexer(code).tokenize();
    const ast = new Parser(tokens).parse();
    const ev = new Evaluator(console.log);
    ev.evaluate(ast);
    return ev;
}

// --- 1. Err.Raise とプロパティ ---
const errCode = `
    Public errNum, errDesc, errSource
    Sub Test()
        On Error Resume Next
        Err.Raise 999, "MySource", "MyDescription"
        errNum = Err.Number
        errDesc = Err.Description
        errSource = Err.Source
        Err.Clear
    End Sub
`;
const ev1 = evalVBA(errCode);
ev1.callProcedure('Test', []);
assert.strictEqual(ev1.env.get('errnum'), 999, 'Err.Number');
assert.strictEqual(ev1.env.get('errdesc'), "MyDescription", 'Err.Description');
assert.strictEqual(ev1.env.get('errsource'), "MySource", 'Err.Source');
assert.strictEqual(ev1.errObj.number, 0, 'Err.Clear 後の状態');
console.log('[PASS] Err.Raise とプロパティ');

// --- 2. 暗黙的なエラー発生時の Err オブジェクト ---
const implicitErrCode = `
    Public errNum2
    Sub Test()
        On Error Resume Next
        Dim x
        x = 1 / 0
        errNum2 = Err.Number
    End Sub
`;
const ev2 = evalVBA(implicitErrCode);
ev2.callProcedure('Test', []);
assert.strictEqual(ev2.env.get('errnum2'), 11, 'Division by zero (11)');
console.log('[PASS] 暗黙的エラー時の Err オブジェクト');

console.log('\n✅ Err Object: 全テスト通過');
