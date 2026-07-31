import { assert, evalVBASingle } from '../../test-libs/test-runner';

const operations = {
    print: 'Print #1, "value"',
    write: 'Write #1, "value"',
    put: 'Put #1, , value',
    get: 'Get #1, , value',
    input: 'Input #1, value',
    lineInput: 'Line Input #1, value',
} as const;

const allowed: Record<keyof typeof operations, string[]> = {
    print: ['Output', 'Append'],
    write: ['Output', 'Append'],
    put: ['Random', 'Binary'],
    get: ['Random', 'Binary'],
    input: ['Input', 'Binary'],
    lineInput: ['Input', 'Binary'],
};

const modes = ['Input', 'Output', 'Append', 'Random', 'Binary'];
let cases = 0;

for (const mode of modes) {
    for (const operation of Object.keys(operations) as (keyof typeof operations)[]) {
        if (allowed[operation].includes(mode)) continue;
        const path = `mode-error-${mode.toLowerCase()}-${operation}.dat`;
        const seed = mode === 'Input' || mode === 'Append'
            ? `        Open "${path}" For Output As #9\n        Close #9\n`
            : '';
        const open = mode === 'Random'
            ? `Open "${path}" For Random As #1 Len = 8`
            : `Open "${path}" For ${mode} As #1`;
        const value = operation === 'put' || operation === 'get'
            ? '        Dim value As Long\n        value = 1\n'
            : '        Dim value As String\n';
        const ev = evalVBASingle(`
            Function ProbeModeError() As Long
${value}${seed}        ${open}
            On Error Resume Next
            ${operations[operation]}
            ProbeModeError = Err.Number
            Close #1
            End Function
        `);
        assert.strictEqual(ev.callProcedure('ProbeModeError', []), 54,
            `${mode}/${operation} must raise Bad file mode (54)`);
        cases++;
    }
}

assert.strictEqual(cases, 18, 'all disallowed file-I/O mode pairs are covered');
console.log(`[PASS] file-I/O mode guards (${cases} disallowed pairs)`);
