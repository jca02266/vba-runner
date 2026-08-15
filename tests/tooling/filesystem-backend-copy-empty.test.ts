import { strict as assert } from 'node:assert';
import * as nodeFs from 'node:fs';
import * as nodeOs from 'node:os';
import * as nodePath from 'node:path';
import { MemoryFileSystem } from '../../src/engine/filesystem';
import { NodeFileSystem } from '../../src/engine/node_filesystem';

type Backend = MemoryFileSystem | NodeFileSystem;

function probe(fs: Backend, root: string) {
    const sourceFile = `${root}/source-file`;
    const sourceFileDestination = `${root}/source-file-destination`;
    fs.writeFileSync(sourceFile, 'not-a-directory');
    fs.mkdirSync(sourceFileDestination, { recursive: true });
    let sourceType = '';
    try {
        fs.copyDirectorySync(sourceFile, sourceFileDestination, { overwrite: false });
    } catch (error) {
        sourceType = (error as { code?: string }).code
            ?? String(error).match(/EEXIST|ENOENT|EINVAL/)?.[0]
            ?? String(error);
    }

    const emptySource = `${root}/empty-source`;
    const emptyDestination = `${root}/empty-destination`;
    fs.mkdirSync(emptySource, { recursive: true });
    fs.copyDirectorySync(emptySource, emptyDestination, { overwrite: false });
    const empty = fs.existsSync(emptyDestination);

    const source = `${root}/source`;
    const destination = `${root}/destination`;
    fs.mkdirSync(`${source}/empty/deep`, { recursive: true });
    fs.writeFileSync(`${source}/file.txt`, 'value');
    fs.mkdirSync(destination, { recursive: true });
    fs.copyDirectorySync(source, destination, { overwrite: false });
    const mixed = [
        fs.existsSync(`${destination}/empty`),
        fs.existsSync(`${destination}/empty/deep`),
        fs.existsSync(`${destination}/file.txt`),
        fs.readFileSync(`${destination}/file.txt`, 'utf8'),
    ];

    const directorySource = `${root}/directory-source`;
    const directoryDestination = `${root}/directory-destination`;
    fs.mkdirSync(`${directorySource}/child`, { recursive: true });
    fs.writeFileSync(`${directorySource}/child/value.txt`, 'new');
    fs.mkdirSync(directoryDestination, { recursive: true });
    fs.writeFileSync(`${directoryDestination}/child`, 'existing');
    let directoryCollision = '';
    try {
        fs.copyDirectorySync(directorySource, directoryDestination, { overwrite: false });
    } catch (error) {
        directoryCollision = (error as { code?: string }).code
            ?? String(error).match(/EEXIST|ENOENT|EINVAL/)?.[0]
            ?? String(error);
    }
    const fileSource = `${root}/file-source`;
    const fileDestination = `${root}/file-destination`;
    fs.mkdirSync(fileSource, { recursive: true });
    fs.writeFileSync(`${fileSource}/child`, 'new');
    fs.mkdirSync(`${fileDestination}/child`, { recursive: true });
    let fileCollision = '';
    try {
        fs.copyDirectorySync(fileSource, fileDestination, { overwrite: true });
    } catch (error) {
        fileCollision = (error as { code?: string }).code
            ?? String(error).match(/EEXIST|ENOENT|EINVAL/)?.[0]
            ?? String(error);
    }

    const overwriteSource = `${root}/overwrite-source`;
    const overwriteDestination = `${root}/overwrite-destination`;
    fs.mkdirSync(overwriteSource, { recursive: true });
    fs.writeFileSync(`${overwriteSource}/value.txt`, 'new');
    fs.mkdirSync(overwriteDestination, { recursive: true });
    fs.writeFileSync(`${overwriteDestination}/value.txt`, 'old');
    let collision = '';
    try {
        fs.copyDirectorySync(overwriteSource, overwriteDestination, { overwrite: false });
    } catch (error) {
        collision = String(error).match(/EEXIST|ENOENT|EINVAL/)?.[0] ?? String(error);
    }
    const collisionState = [
        collision,
        fs.readFileSync(`${overwriteDestination}/value.txt`, 'utf8'),
        fs.existsSync(`${overwriteSource}/value.txt`),
    ];
    fs.copyDirectorySync(overwriteSource, overwriteDestination, { overwrite: true });
    const overwriteState = [fs.readFileSync(`${overwriteDestination}/value.txt`, 'utf8')];
    return { sourceType, empty, mixed, directoryCollision, fileCollision, collisionState, overwriteState };
}

const memory = probe(new MemoryFileSystem(), '/copy-empty-parity');
const nodeRoot = nodeFs.mkdtempSync(nodePath.join(nodeOs.tmpdir(), 'vba-copy-empty-'));
try {
    const node = probe(new NodeFileSystem(), nodeRoot);
    assert.equal(memory.sourceType, 'ENOENT');
    assert.equal(node.sourceType, 'ENOENT');
    assert.equal(memory.directoryCollision, 'EEXIST');
    assert.equal(node.directoryCollision, 'EEXIST');
    assert.equal(memory.fileCollision, 'EEXIST');
    assert.equal(node.fileCollision, 'EEXIST');
    console.log(`Memory ${JSON.stringify(memory)}`);
    console.log(`Node ${JSON.stringify(node)}`);
} finally {
    nodeFs.rmSync(nodeRoot, { recursive: true, force: true });
}
