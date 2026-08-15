import { strict as assert } from 'node:assert';
import * as nodeFs from 'node:fs';
import * as nodeOs from 'node:os';
import * as nodePath from 'node:path';
import { evalVBASingle } from '../../test-libs/test-runner';
import { MemoryFileSystem } from '../../src/engine/filesystem';
import { NodeFileSystem } from '../../src/engine/node_filesystem';

const source = String.raw`
Public copyErr
Sub Probe()
    Dim fso As Object, stream As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    fso.CreateFolder "source"
    fso.CreateFolder "source\child"
    fso.CreateFolder "destination"
    Set stream = fso.CreateTextFile("destination\child")
    stream.Write "existing": stream.Close
    On Error Resume Next
    Err.Clear: fso.CopyFolder "source", "destination", True: copyErr = Err.Number
    On Error GoTo 0
End Sub
`;

function run(fs: MemoryFileSystem | NodeFileSystem, sandboxRoot?: string) {
    const ev = evalVBASingle(source, { fs, ...(sandboxRoot ? { sandboxRoot } : {}) });
    ev.callProcedure('Probe', []);
    return ev.env.get('copyerr');
}

const memory = run(new MemoryFileSystem());
assert.equal(memory, 58);
const root = nodeFs.mkdtempSync(nodePath.join(nodeOs.tmpdir(), 'vba-fso-copy-type-'));
try {
    const node = run(new NodeFileSystem(), root);
    assert.equal(node, memory);
    console.log(`Memory COPY-TYPE=${memory}`);
    console.log(`Node COPY-TYPE=${node}`);
} finally {
    nodeFs.rmSync(root, { recursive: true, force: true });
}
