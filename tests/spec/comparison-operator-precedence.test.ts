import { evalVBASingle, assert } from '../../test-libs/test-runner';

const ev = evalVBASingle(`
    Function TestComparisonOrder() As String
        TestComparisonOrder = CStr(1 = 2 < 3)
    End Function
`);

assert.strictEqual(ev.callProcedure('TestComparisonOrder', []), 'True');

console.log('✅ Comparison operators associate left-to-right');
