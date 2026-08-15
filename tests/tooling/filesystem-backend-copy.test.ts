import { strict as assert } from 'node:assert';
import { MemoryFileSystem } from '../../src/engine/filesystem';
import { NodeFileSystem } from '../../src/engine/node_filesystem';
import * as nodeFs from 'node:fs';
import * as nodePath from 'node:path';
import * as nodeOs from 'node:os';

type Backend = MemoryFileSystem | NodeFileSystem;

// FileSystem's documented preflight contract: a collision with overwrite=false
// must not leave an earlier source entry partially copied.
function assertCopyDirectoryCollision(fs: Backend, root: string): void {
    fs.mkdirSync(`${root}/source`, { recursive: true });
    fs.writeFileSync(`${root}/source/first.txt`, 'first');
    fs.writeFileSync(`${root}/source/second.txt`, 'second');
    fs.mkdirSync(`${root}/dest`, { recursive: true });
    fs.writeFileSync(`${root}/dest/second.txt`, 'existing');
    assert.throws(() => fs.copyDirectorySync(`${root}/source`, `${root}/dest`), /exists|EEXIST/);
    assert.equal(fs.existsSync(`${root}/dest/first.txt`), false,
        'CopyDirectory collision leaves no partial first file');
}

assertCopyDirectoryCollision(new MemoryFileSystem(), '/copy-collision');

const nodeRoot = nodeFs.mkdtempSync(nodePath.join(nodeOs.tmpdir(), 'vba-copy-collision-'));
try {
    assertCopyDirectoryCollision(new NodeFileSystem(), nodeRoot);
} finally {
    nodeFs.rmSync(nodeRoot, { recursive: true, force: true });
}

console.log('[PASS] Memory/Node CopyDirectory collision atomicity');
