import { evalVBASingle, assert } from '../../test-libs/test-runner';

const probe = evalVBASingle(`
    Function ProbeNonFinite() As String
        Dim result As String, value As Variant
        On Error Resume Next
        value = Rnd(-1E+308)
        result = result & Err.Number & ","
        Err.Clear
        value = Rnd()
        result = result & Err.Number & "," & CStr(value) & ","
        Err.Clear
        value = FV(1E+308, 1E+308, 0)
        result = result & Err.Number & ","
        Err.Clear
        value = TimeSerial(-1E+308, &HFFFFFFFF, 1E+308)
        result = result & Err.Number & ","
        Err.Clear
        value = SLN(1E+308, -1E+308, 1E-308)
        result = result & Err.Number & ","
        Err.Clear
        value = NPV(1E-308, Array(1E+308, 1E+308))
        result = result & Err.Number & ","
        Err.Clear
        value = DateSerial(1E+308, 1, 1)
        result = result & Err.Number & ","
        Err.Clear
        value = DateAdd("d", 1E+308, DateSerial(2024, 1, 1))
        result = result & Err.Number & ","
        Err.Clear
        value = DateSerial(10000, 1, 1)
        result = result & Err.Number & ","
        Err.Clear
        value = DateAdd("d", 1, DateSerial(9999, 12, 31))
        result = result & Err.Number
        ProbeNonFinite = result
    End Function
`);

assert.strictEqual(probe.callProcedure('ProbeNonFinite', []), '6,0,0.000614211312495172,6,6,6,6,6,6,6,6',
    'non-finite random, financial, and date results become VBA errors');
console.log('[PASS] non-finite builtin boundaries');
