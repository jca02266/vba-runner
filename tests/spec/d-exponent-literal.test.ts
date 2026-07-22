import { evalVBASingle, assert } from '../../test-libs/test-runner';

const ev = evalVBASingle(`
    Public largeValue As Double, smallValue As Double, lowerCaseValue As Double
    Public largeType As String, smallType As String, lowerCaseType As String

    Sub TestDExponentLiterals()
        largeValue = 1.5D+10
        smallValue = 2D5
        lowerCaseValue = 1d-2
        largeType = TypeName(1.5D+10)
        smallType = TypeName(2D5)
        lowerCaseType = TypeName(1d-2)
    End Sub
`);

ev.callProcedure('TestDExponentLiterals', []);

assert.strictEqual(ev.env.get('largeValue'), 15000000000);
assert.strictEqual(ev.env.get('smallValue'), 200000);
assert.strictEqual(ev.env.get('lowerCaseValue'), 0.01);
assert.deepStrictEqual([
    ev.env.get('largeType'), ev.env.get('smallType'), ev.env.get('lowerCaseType'),
], ['Double', 'Double', 'Double']);

console.log('✅ VBA Double D/d exponent literals parse and retain their subtype');
