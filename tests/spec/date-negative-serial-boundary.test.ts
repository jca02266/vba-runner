import { strict as assert } from 'node:assert';
import { evalVBASingle } from '../../test-libs/test-runner';

const source = String.raw`Option Explicit
Public Function Probe() As String
    Probe = CStr(DatePart("yyyy", CDate(-1))) & "|" & _
        CStr(DatePart("m", CDate(-1))) & "|" & _
        CStr(DatePart("d", CDate(-1))) & "|" & _
        CStr(DateDiff("d", CDate(0), CDate(-1))) & "|" & _
        CStr(Hour(CDate(-0.5))) & "|" & _
        CStr(Minute(CDate(-0.5))) & "|" & _
        CStr(Second(CDate(-0.00001)))
End Function`;

const ev = evalVBASingle(source);
assert.equal(ev.callProcedure('Probe', []), '1899|12|29|-1|12|0|0');
console.log('[PASS] negative date serial component boundary (7 cases)');
