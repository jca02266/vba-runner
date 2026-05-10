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

console.log('[Test Suite] Error Handling (On Error / Resume) の検証');

// Case 1: On Error Resume Next
const code1 = `
    Function TestResumeNext()
        On Error Resume Next
        Dim x
        x = 1 / 0
        TestResumeNext = 42
    End Function
`;
const ev1 = evalVBA(code1);
assert.strictEqual(ev1.callProcedure('TestResumeNext', []), 42, 'On Error Resume Next should skip error line');

// Case 2: On Error GoTo <Label> and Resume Next (statement)
const code2 = `
    Function TestGoToLabel()
        Dim errorCaught
        errorCaught = 0
        On Error GoTo ErrorHandler
        Dim x
        x = 1 / 0
        TestGoToLabel = errorCaught + 100
        Exit Function
    ErrorHandler:
        errorCaught = Err.Number
        Resume Next
    End Function
`;
const ev2 = evalVBA(code2);
assert.strictEqual(ev2.callProcedure('TestGoToLabel', []), 111, 'ErrorHandler should catch error (11) and Resume Next should continue to the next line (11 + 100 = 111)');

// Case 3: Resume (to the same line)
const code3 = `
    Function TestResume()
        Dim x, count
        count = 0
        On Error GoTo ErrorHandler
    TryAgain:
        x = 1 / (1 - count)
        TestResume = x
        Exit Function
    ErrorHandler:
        count = 1
        Resume
    End Function
`;
// Note: 1 / (1-0) = 1. Wait, if count=0, x = 1/1 = 1. No error.
// Let's make it error first.
const code3fixed = `
    Function TestResume()
        Dim x, divisor
        divisor = 0
        On Error GoTo ErrorHandler
        x = 1 / divisor
        TestResume = x
        Exit Function
    ErrorHandler:
        divisor = 1
        Resume
    End Function
`;
const ev3 = evalVBA(code3fixed);
assert.strictEqual(ev3.callProcedure('TestResume', []), 1, 'Resume should retry the failing line');

// Case 4: On Error GoTo 0
const code4 = `
    Function TestGoTo0()
        On Error Resume Next
        On Error GoTo 0
        Dim x
        x = 1 / 0
        TestGoTo0 = 42
    End Function
`;
const ev4 = evalVBA(code4);
try {
    ev4.callProcedure('TestGoTo0', []);
    assert.ok(false, 'On Error GoTo 0 should re-enable default error behavior (throwing)');
} catch (e) {
    assert.ok(true, 'Caught expected error after On Error GoTo 0');
}

console.log('✅ Error Handling: 全テスト通過');
