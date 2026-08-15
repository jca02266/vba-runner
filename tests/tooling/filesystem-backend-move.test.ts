import { strict as assert } from 'node:assert';
import * as nodeFs from 'node:fs';
import * as nodeOs from 'node:os';
import * as nodePath from 'node:path';
import { MemoryFileSystem } from '../../src/engine/filesystem';
import { NodeFileSystem } from '../../src/engine/node_filesystem';

type Backend = MemoryFileSystem | NodeFileSystem;

function probe(fs: Backend, root: string): string {
  fs.mkdirSync(`${root}/source`, { recursive: true });
  fs.mkdirSync(`${root}/destination`, { recursive: true });
  let result = 'ok';
  try {
    fs.moveDirectorySync(`${root}/source`, `${root}/destination`);
  } catch (error) {
    result = String((error as { code?: string }).code ?? error).includes('EEXIST')
      ? 'EEXIST' : String(error);
  }
  return `${result}:source=${fs.existsSync(`${root}/source`)}:destination=${fs.existsSync(`${root}/destination`)}`;
}

const memoryResult = probe(new MemoryFileSystem(), '/move-collision');
const nodeRoot = nodeFs.mkdtempSync(nodePath.join(nodeOs.tmpdir(), 'vba-move-collision-'));
let nodeResult: string;
try {
  nodeResult = probe(new NodeFileSystem(), nodeRoot);
} finally {
  nodeFs.rmSync(nodeRoot, { recursive: true, force: true });
}

assert.equal(memoryResult, 'EEXIST:source=true:destination=true');
assert.equal(nodeResult, 'ok:source=false:destination=true');
console.log(`Memory empty-destination=${memoryResult}`);
console.log(`Node empty-destination=${nodeResult}`);
