import { assert, evalVBASingle } from '../../test-libs/test-runner';

const ev = evalVBASingle(`
    Function ProbeAppendHandles() As Long
        Dim result As Long
        Open "append-multi.dat" For Output As #9
        Print #9, "seed"
        Close #9
        Open "append-multi.dat" For Append Access Write Shared As #1
        Open "append-multi.dat" For Append Access Write Shared As #2
        Print #1, "A"
        Print #2, "B"
        result = LOF(1) + Seek(1) + Seek(2)
        Close #1
        Close #2
        ProbeAppendHandles = result
    End Function
`);

assert.strictEqual(ev.callProcedure('ProbeAppendHandles', []), 35,
    'each Append handle advances from the physical end');
const content = (ev as any).fs.readFileSync('/sandbox/append-multi.dat').toString();
assert.strictEqual(content, 'seed\r\nA\r\nB\r\n',
    'shared Append handles never overwrite an earlier append');
console.log('[PASS] shared Append handles preserve physical append order');
