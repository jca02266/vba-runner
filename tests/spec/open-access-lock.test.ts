import { evalVBASingle, assert } from '../../test-libs/test-runner';
import { MemoryFileSystem } from '../../src/engine/filesystem';

function run(code: string): any {
    const ev = evalVBASingle(code, { fs: new MemoryFileSystem() });
    return ev.callProcedure('Test', []);
}

function assertOpenResult(first: string, second: string, expectedError?: number): void {
    const code = `
        Sub Test()
            Open "lock.dat" For Binary Access Write As #9
            Put #9, , CByte(1)
            Close #9
            Open "lock.dat" For Binary ${first} As #1
            Open "lock.dat" For Binary ${second} As #2
            Close #2
            Close #1
        End Sub
    `;
    let error: any;
    try { run(code); } catch (e) { error = e; }
    assert.strictEqual(error?.number, expectedError,
        `${first} -> ${second} should ${expectedError === undefined ? 'succeed' : `raise ${expectedError}`}`);
}

assertOpenResult('Access Read Lock Read', 'Access Read Lock Read', 70);
assertOpenResult('Access Read Lock Read', 'Access Write Shared');
assertOpenResult('Access Read Lock Write', 'Access Read Shared');
assertOpenResult('Access Read Lock Write', 'Access Write Shared', 70);
assertOpenResult('Access Read Write Shared', 'Access Read Write Shared');

console.log('✅ Open Access/Lock compatibility follows the Excel probe matrix');
