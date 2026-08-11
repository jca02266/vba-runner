import { strict as assert } from 'node:assert';
import { evalVBASingle } from '../../test-libs/test-runner';

const ev = evalVBASingle(String.raw`
Option Explicit
Public Function ProbeRadixConversions() As String
    ProbeRadixConversions = CStr(CLngLng("&HFFFFFFFFFFFFFFFF")) & "|" & _
        CStr(CCur("&HFFFFFFFF")) & "|" & _
        CStr(Val("&HFFFFFFFFFFFFFFFF")) & "|" & _
        CStr(Val("&O1777777777777777777777"))
End Function
`);

assert.equal(ev.callProcedure('ProbeRadixConversions', []), '-1|-1|-1|-1');
console.log('[PASS] radix conversion width and sign regressions');
