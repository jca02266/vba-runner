import { strict as assert } from 'node:assert';
import { evalVBASingle } from '../../test-libs/test-runner';

const cases = [
    ['FV(0.01,12,-100,0,Null)', '', 94],
    ['FV(0.01,12,-100,0,Empty)', '', 0],
    ['PMT(0.01,12,1000,Null)', '', 94],
    ['NPer(0.01,Null,1000,0)', '', 94],
    ['Rate(12,-100,100,Null)', '', 94],
    ['MIRR(flows,Null,0.1)', 'Dim flows(0 To 1) As Double: flows(0) = -100: flows(1) = 100', 94],
    ['NPV(Null,flows)', 'Dim flows(0 To 1) As Double: flows(0) = -100: flows(1) = 100', 94],
] as const;

for (const [expression, setup, expected] of cases) {
    const evaluator = evalVBASingle(String.raw`Function Probe() As Long
        On Error Resume Next
        ${setup}
        Dim value As Variant
        value = ${expression}
        Probe = Err.Number
    End Function`);
    assert.equal(evaluator.callProcedure('Probe', []), expected, expression);
}

console.log(`[PASS] financial Optional/Null matrix (${cases.length} boundaries)`);
