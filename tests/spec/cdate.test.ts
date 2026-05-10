import { assert } from '../ts/test-runner';
import { Evaluator, VbaDate } from '../../src/compiler/evaluator';
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

const code = `
    Function TestCDateString()
        TestCDateString = CDate("2025/05/10")
    End Function

    Function TestCDateNumber()
        TestCDateNumber = CDate(45658)
    End Function

    Function TestCVDateNull()
        TestCVDateNull = CVDate(Null)
    End Function

    Function TestCVDateDate()
        TestCVDateDate = CVDate(#2025/05/10#)
    End Function
`;

const ev = evalVBA(code);

console.log('[Test Suite] CDate / CVDate の検証');

const dateStr = ev.callProcedure('TestCDateString', []);
assert.ok(dateStr instanceof VbaDate, 'CDate(string) should return VbaDate');
assert.strictEqual(ev.callProcedure('TypeName', [dateStr]), 'Date', 'TypeName should be Date');

const dateNum = ev.callProcedure('TestCDateNumber', []);
assert.ok(dateNum instanceof VbaDate, 'CDate(number) should return VbaDate');
assert.strictEqual(ev.callProcedure('TypeName', [dateNum]), 'Date', 'TypeName should be Date');

const vdateNull = ev.callProcedure('TestCVDateNull', []);
assert.strictEqual(ev.callProcedure('TypeName', [vdateNull]), 'Null', 'CVDate(Null) should be Null');

const vdateDate = ev.callProcedure('TestCVDateDate', []);
assert.ok(vdateDate instanceof VbaDate, 'CVDate(Date) should return VbaDate');

console.log('✅ CDate / CVDate: 全テスト通過');
