import { strict as assert } from 'node:assert';
import { evalVBASingle } from '../../test-libs/test-runner';

const cases = [
    ['Rate(0,-100,100)', '', 5],
    ['NPer(0,0,0,0)', '', 5],
    ['MIRR(flows,0.1,-1)', 'Dim flows(0 To 1) As Double: flows(0) = -100: flows(1) = 100', 5],
    ['IPmt(0.01,0,12,1000)', '', 5],
    ['PPmt(0.01,13,12,1000)', '', 5],
    ['DDB(1000,100,5,0)', '', 5],
    ['SYD(1000,100,5,6)', '', 5],
    ['Pmt(0.01,0,1000)', '', 5],
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

console.log(`[PASS] financial invalid-argument matrix (${cases.length} boundaries)`);
