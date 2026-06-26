import { evalVBASingle, assert, vbaTrue, vbaFalse } from '../../test-libs/test-runner';

function evalVBA(code: string): any {
    return evalVBASingle(code);
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

assert.strictEqual(ev.callProcedure('TestDateLiteral', []), vbaTrue, 'IsDate(#2025/01/01#) -> True');
assert.strictEqual(ev.callProcedure('GetDateTypeName', []), 'Date', 'TypeName(#2025/01/01#) -> Date');
assert.strictEqual(ev.callProcedure('TestDateArithmetic', []), 'Date', 'TypeName(#2025/01/01# + 1) -> Date');
assert.strictEqual(ev.callProcedure('TestDateString', []), vbaTrue, 'IsDate("2025/01/01") -> True');
assert.strictEqual(ev.callProcedure('TestInvalidString', []), vbaFalse, 'IsDate("not a date") -> False');
assert.strictEqual(ev.callProcedure('TestEmpty', []), vbaFalse, 'IsDate(Empty) -> False');
assert.strictEqual(ev.callProcedure('TestNull', []), vbaFalse, 'IsDate(Null) -> False');

console.log('[PASS] IsDate の基本動作');

console.log('\n✅ IsDate: 全テスト通過');
