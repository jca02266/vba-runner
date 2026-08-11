import { strict as assert } from 'node:assert';
import { evalVBASingle } from '../../test-libs/test-runner';

type Case = { base: 'decimal' | 'hex' | 'octal'; literal: string; expected: string };

const cases: Case[] = [
    { base: 'decimal', literal: '0^', expected: '0:0' },
    { base: 'decimal', literal: '1^', expected: '0:1' },
    { base: 'decimal', literal: '-1^', expected: '0:-1' },
    { base: 'decimal', literal: '9223372036854775807^', expected: '0:9223372036854775807' },
    { base: 'decimal', literal: '-9223372036854775808^', expected: '0:-9223372036854775808' },
    { base: 'decimal', literal: '9223372036854775808^', expected: '6:0' },
    { base: 'hex', literal: '&H0^', expected: '0:0' },
    { base: 'hex', literal: '&H7FFFFFFFFFFFFFFF^', expected: '0:9223372036854775807' },
    { base: 'hex', literal: '&H8000000000000000^', expected: '0:-9223372036854775808' },
    { base: 'hex', literal: '&HFFFFFFFFFFFFFFFF^', expected: '0:-1' },
    { base: 'hex', literal: '&H20000000000001^', expected: '0:9007199254740993' },
    { base: 'octal', literal: '&O0^', expected: '0:0' },
    { base: 'octal', literal: '&O777777777777777777777^', expected: '0:9223372036854775807' },
    { base: 'octal', literal: '&O1000000000000000000000^', expected: '0:-9223372036854775808' },
    { base: 'octal', literal: '&O400000000000000000001^', expected: '0:4611686018427387905' },
];

for (const testCase of cases) {
    const ev = evalVBASingle(String.raw`Option Explicit
Public Function Probe() As String
    Dim value As LongLong, errNo As Long
    On Error Resume Next
    Err.Clear
    value = ${testCase.literal}
    errNo = Err.Number
    Probe = CStr(errNo) & ":" & CStr(value)
End Function`);
    assert.equal(ev.callProcedure('Probe', []), testCase.expected,
        `${testCase.base} LongLong literal ${testCase.literal}`);
}

assert.equal(cases.length, 15);
console.log(`[PASS] LongLong literal boundary matrix (${cases.length} cases)`);
