import { evalVBASingle, assert } from '../../test-libs/test-runner';

const ev = evalVBASingle(`
    Function TestSharedOpen() As String
        Open "shared.dat" For Random Access Read Shared As #1
        Close #1
        TestSharedOpen = "ok"
    End Function
`);

assert.strictEqual(ev.callProcedure('TestSharedOpen', []), 'ok');

console.log('✅ Open accepts the standalone Shared lock clause');
