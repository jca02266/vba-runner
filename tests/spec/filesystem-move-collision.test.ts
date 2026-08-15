import { strict as assert } from 'node:assert';
import * as nodeFs from 'node:fs';
import * as nodeOs from 'node:os';
import * as nodePath from 'node:path';
import { evalVBASingle } from '../../test-libs/test-runner';
import { MemoryFileSystem } from '../../src/engine/filesystem';
import { NodeFileSystem } from '../../src/engine/node_filesystem';

const source = String.raw`
Public moveFileErr, fileMoveErr, moveFolderErr, folderMoveErr
Sub Probe()
    Dim fso As Object, stream As Object, file As Object, folder As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    fso.CreateFolder "folder-source"
    fso.CreateFolder "folder-destination"
    Set stream = fso.CreateTextFile("file-source.txt")
    stream.Write "source": stream.Close
    Set stream = fso.CreateTextFile("file-destination.txt")
    stream.Write "destination": stream.Close

    On Error Resume Next
    Err.Clear: fso.MoveFile "file-source.txt", "file-destination.txt": moveFileErr = Err.Number
    Err.Clear: Set file = fso.GetFile("file-source.txt"): file.Move "file-destination.txt": fileMoveErr = Err.Number
    Err.Clear: fso.MoveFolder "folder-source", "folder-destination": moveFolderErr = Err.Number
    Err.Clear: Set folder = fso.GetFolder("folder-source"): folder.Move "folder-destination": folderMoveErr = Err.Number
    On Error GoTo 0
End Sub
`;

function run(fs: MemoryFileSystem | NodeFileSystem, sandboxRoot?: string) {
    const ev = evalVBASingle(source, { fs, ...(sandboxRoot ? { sandboxRoot } : {}) });
    ev.callProcedure('Probe', []);
    return {
        moveFile: ev.env.get('movefileerr'),
        fileMove: ev.env.get('filemoveerr'),
        moveFolder: ev.env.get('movefoldererr'),
        folderMove: ev.env.get('foldermoveerr'),
    };
}

const memory = run(new MemoryFileSystem());
assert.deepEqual(memory, { moveFile: 58, fileMove: 58, moveFolder: 58, folderMove: 58 });
const root = nodeFs.mkdtempSync(nodePath.join(nodeOs.tmpdir(), 'vba-fso-move-collision-'));
try {
    const node = run(new NodeFileSystem(), root);
    assert.deepEqual(node, memory);
    console.log(`Memory ${JSON.stringify(memory)}`);
    console.log(`Node ${JSON.stringify(node)}`);
} finally {
    nodeFs.rmSync(root, { recursive: true, force: true });
}
