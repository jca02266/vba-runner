import { strict as assert } from 'node:assert';
import { evalVBASingle } from '../../test-libs/test-runner';

const ev = evalVBASingle(String.raw`
Option Explicit
Public Function ProbeRadixConversions() As String
    ProbeRadixConversions = CStr(CLngLng("&HFFFFFFFFFFFFFFFF")) & "|" & _
        CStr(CCur("&HFFFFFFFF")) & "|" & _
        CStr(Val("&HFFFFFFFFFFFFFFFF")) & "|" & _
        CStr(Val("&O37777777777")) & "|" & _
        CStr(CCur("&O37777777777")) & "|" & CStr(CDec("&HFFFFFFFF"))
End Function
Public Function ProbeExplicitSignError() As Long
    On Error Resume Next
    Dim value As Variant
    value = CDec("+&H10")
    ProbeExplicitSignError = Err.Number
End Function
`);

assert.equal(ev.callProcedure('ProbeRadixConversions', []), '-1|-1|-1|-1|-1|-1');
assert.equal(ev.callProcedure('ProbeExplicitSignError', []), 13);
console.log('[PASS] radix conversion width and sign regressions');
