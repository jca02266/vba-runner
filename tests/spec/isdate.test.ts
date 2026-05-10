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

console.log('[Test Suite] IsDate の基本動作');

const code = `
    Function TestDateLiteral()
        TestDateLiteral = IsDate(#2025/01/01#)
    End Function

    Function GetDateTypeName()
        GetDateTypeName = TypeName(#2025/01/01#)
    End Function

    Function TestDateArithmetic()
        TestDateArithmetic = TypeName(#2025/01/01# + 1)
    End Function

    Function TestDateString()
        TestDateString = IsDate("2025/01/01")
    End Function

    Function TestInvalidString()
        TestInvalidString = IsDate("not a date")
    End Function

    Function TestEmpty()
        Dim v
        TestEmpty = IsDate(v)
    End Function

    Function TestNull()
        TestNull = IsDate(Null)
    End Function
`;

const ev = evalVBA(code);

assert.isTrue(ev.callProcedure('TestDateLiteral', []), 'IsDate(#2025/01/01#) -> True');
assert.strictEqual(ev.callProcedure('GetDateTypeName', []), 'Date', 'TypeName(#2025/01/01#) -> Date');
assert.strictEqual(ev.callProcedure('TestDateArithmetic', []), 'Date', 'TypeName(#2025/01/01# + 1) -> Date');
assert.isTrue(ev.callProcedure('TestDateString', []), 'IsDate("2025/01/01") -> True');
assert.isFalse(ev.callProcedure('TestInvalidString', []), 'IsDate("not a date") -> False');
assert.isFalse(ev.callProcedure('TestEmpty', []), 'IsDate(Empty) -> False');
assert.isFalse(ev.callProcedure('TestNull', []), 'IsDate(Null) -> False');

console.log('[PASS] IsDate の基本動作');

console.log('\n✅ IsDate: 全テスト通過');
