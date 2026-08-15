import { strict as assert } from 'node:assert';
import * as nodeFs from 'node:fs';
import * as nodeOs from 'node:os';
import * as nodePath from 'node:path';
import { MemoryFileSystem } from '../../src/engine/filesystem';
import { NodeFileSystem } from '../../src/engine/node_filesystem';

type Backend = MemoryFileSystem | NodeFileSystem;

function probe(fs: Backend, root: string): string {
    const source = `${root}/source`;
    const destination = `${root}/destination`;
    fs.mkdirSync(`${source}/a`, { recursive: true });
    fs.mkdirSync(`${source}/z`, { recursive: true });
    fs.writeFileSync(`${source}/a/first.txt`, 'first');
    fs.writeFileSync(`${source}/z/conflict.txt`, 'new');
    fs.mkdirSync(`${destination}/z`, { recursive: true });
    fs.writeFileSync(`${destination}/z/conflict.txt`, 'old');
    let error = '';
    try {
        fs.copyDirectorySync(source, destination, { overwrite: false });
    } catch (caught) {
        error = (caught as { code?: string }).code ?? String(caught).match(/EEXIST|ENOENT|EINVAL/)?.[0] ?? String(caught);
    }
    return `${error}:${fs.existsSync(`${destination}/a/first.txt`)}:${fs.readFileSync(`${destination}/z/conflict.txt`, 'utf8')}:${fs.existsSync(`${source}/a/first.txt`)}`;
}

const memory = probe(new MemoryFileSystem(), '/copy-nested');
const root = nodeFs.mkdtempSync(nodePath.join(nodeOs.tmpdir(), 'vba-copy-nested-'));
try {
    const node = probe(new NodeFileSystem(), root);
    assert.equal(memory, 'EEXIST:false:old:true');
    assert.equal(node, memory);
    console.log(`Memory NESTED=${memory}`);
    console.log(`Node NESTED=${node}`);
} finally {
    nodeFs.rmSync(root, { recursive: true, force: true });
}
