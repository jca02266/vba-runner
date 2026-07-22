import { evalVBASingle, assert } from '../../test-libs/test-runner';

const ev = evalVBASingle(`
    Public lowerBound As Long, upperBound As Long
    Public b0 As Byte, b1 As Byte, b2 As Byte, b3 As Byte

    Sub TestAssignment()
        Dim b() As Byte
        b = "Aあ"
        lowerBound = LBound(b)
        upperBound = UBound(b)
        b0 = b(0): b1 = b(1): b2 = b(2): b3 = b(3)
    End Sub
`);

ev.callProcedure('TestAssignment', []);

assert.strictEqual(ev.env.get('lowerBound'), 0);
assert.strictEqual(ev.env.get('upperBound'), 3);
assert.deepStrictEqual([
    ev.env.get('b0'), ev.env.get('b1'), ev.env.get('b2'), ev.env.get('b3'),
], [0x41, 0x00, 0x42, 0x30]);

console.log('✅ String assignment to Byte() uses UTF-16LE code units');
