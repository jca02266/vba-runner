/**
 * Exit Statement (§5.4.2.17 & §5.4.2.18) のテスト
 */
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

function runFunc(code: string, name: string, args: any[] = []): any {
    return evalVBA(code).callProcedure(name, args);
}

// --- 1. Exit Sub ---
const exitSubCode = `
    Dim result
    Sub TestExitSub(n)
        result = "Started"
        If n > 0 Then
            Exit Sub
        End If
        result = "Finished"
    End Sub
`;
const ev1 = evalVBA(exitSubCode);
ev1.callProcedure('TestExitSub', [1]);
assert.strictEqual(ev1.env.get('result'), "Started", 'Exit Sub で中断される');
ev1.callProcedure('TestExitSub', [0]);
assert.strictEqual(ev1.env.get('result'), "Finished", 'Exit Sub を通らなければ最後まで実行される');
console.log('[PASS] Exit Sub');

// --- 2. Exit Function ---
const exitFuncCode = `
    Function TestExitFunc(n)
        TestExitFunc = "Started"
        If n > 0 Then
            TestExitFunc = "Exited"
            Exit Function
        End If
        TestExitFunc = "Finished"
    End Function
`;
assert.strictEqual(runFunc(exitFuncCode, 'TestExitFunc', [1]), "Exited", 'Exit Function で戻り値を返して中断');
assert.strictEqual(runFunc(exitFuncCode, 'TestExitFunc', [0]), "Finished", 'Exit Function を通らなければ最後まで実行');
console.log('[PASS] Exit Function');

// --- 3. ループ内からの Exit Sub ---
const loopExitSubCode = `
    Dim count
    Sub LoopExit(n)
        count = 0
        Dim i
        For i = 1 To 10
            count = count + 1
            If count = n Then Exit Sub
        Next
    End Sub
`;
const ev3 = evalVBA(loopExitSubCode);
ev3.callProcedure('LoopExit', [5]);
assert.strictEqual(ev3.env.get('count'), 5, 'ループ内から Exit Sub');
console.log('[PASS] ループ内からの Exit Sub');

console.log('\n✅ Exit Sub & Function: 全テスト通過');
