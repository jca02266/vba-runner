import { strict as assert } from 'node:assert';
import * as nodeFs from 'node:fs';
import * as nodeOs from 'node:os';
import * as nodePath from 'node:path';
import { MemoryFileSystem } from '../../src/engine/filesystem';
import { NodeFileSystem } from '../../src/engine/node_filesystem';

type Backend = MemoryFileSystem | NodeFileSystem;

function write(fs: Backend, file: string, text: string): void {
  const fd = fs.openSync(file, 'w');
  try { fs.writeSync(fd, text); } finally { fs.closeSync(fd); }
}

function probe(fs: Backend, root: string): string {
  fs.mkdirSync(root, { recursive: true });
  write(fs, `${root}/source.txt`, 'source');
  write(fs, `${root}/destination.txt`, 'destination');
  let result = 'ok';
  try {
    fs.moveFileSync(`${root}/source.txt`, `${root}/destination.txt`);
  } catch (error) {
    result = String((error as { code?: string }).code ?? error).includes('EEXIST')
      ? 'EEXIST' : String(error);
  }
  return `${result}:source=${fs.existsSync(`${root}/source.txt`)}:destination=${fs.existsSync(`${root}/destination.txt`)}`;
}

const memoryResult = probe(new MemoryFileSystem(), '/move-file-collision');
const nodeRoot = nodeFs.mkdtempSync(nodePath.join(nodeOs.tmpdir(), 'vba-move-file-'));
let nodeResult: string;
try {
  nodeResult = probe(new NodeFileSystem(), nodeRoot);
} finally {
  nodeFs.rmSync(nodeRoot, { recursive: true, force: true });
}

assert.equal(memoryResult, 'EEXIST:source=true:destination=true');
assert.equal(nodeResult, 'EEXIST:source=true:destination=true');
console.log(`Memory existing-file=${memoryResult}`);
console.log(`Node existing-file=${nodeResult}`);
