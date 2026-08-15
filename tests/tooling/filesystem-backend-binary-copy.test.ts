import { strict as assert } from 'node:assert';
import * as nodeFs from 'node:fs';
import * as nodeOs from 'node:os';
import * as nodePath from 'node:path';
import { MemoryFileSystem } from '../../src/engine/filesystem';
import { NodeFileSystem } from '../../src/engine/node_filesystem';

type Backend = MemoryFileSystem | NodeFileSystem;

function writeBytes(fs: Backend, file: string, bytes: Uint8Array): void {
  const fd = fs.openSync(file, 'w');
  try {
    fs.writeSync(fd, bytes, 0, bytes.length, 0);
  } finally {
    fs.closeSync(fd);
  }
}

function readBytes(fs: Backend, file: string, length: number): number[] {
  const fd = fs.openSync(file, 'r');
  try {
    const bytes = new Uint8Array(length);
    const read = fs.readSync(fd, bytes, 0, length, 0);
    return [...bytes.slice(0, read)];
  } finally {
    fs.closeSync(fd);
  }
}

function probe(fs: Backend, root: string, label: string): string[] {
  fs.mkdirSync(root, { recursive: true });
  const source = `${root}/source.bin`;
  const destination = `${root}/destination.bin`;
  const cases: Array<[string, number[]]> = [
    ['empty', []],
    ['zero', [0x00]],
    ['ff', [0xff]],
    ['utf8', [0xe3, 0x81, 0x82, 0x41]],
  ];
  return cases.map(([name, values]) => {
    fs.rmSync(destination, { force: true });
    const input = new Uint8Array(values);
    writeBytes(fs, source, input);
    input[0] = 0x7f;
    const original = readBytes(fs, source, values.length);
    fs.copyFileSync(source, destination);
    const copied = readBytes(fs, destination, values.length);
    assert.deepEqual(copied, values, `${label} ${name} copy`);
    assert.deepEqual(original, values, `${label} ${name} input ownership`);
    let noOverwrite = 'missing';
    try {
      fs.copyFileSync(source, destination);
    } catch (error) {
      noOverwrite = String((error as { code?: string }).code ?? error).includes('EEXIST')
        ? 'ok' : String(error);
    }
    assert.equal(noOverwrite, 'ok', `${label} ${name} overwrite=false`);
    fs.copyFileSync(source, destination, { overwrite: true });
    assert.deepEqual(readBytes(fs, destination, values.length), values,
      `${label} ${name} overwrite=true`);
    return `${label} ${name} bytes=${values.map((v) => v.toString(16).padStart(2, '0')).join(' ')} overwrite=false=ok overwrite=true=ok input-owner=ok`;
  });
}

const memory = probe(new MemoryFileSystem(), '/binary-copy', 'Memory');
const root = nodeFs.mkdtempSync(nodePath.join(nodeOs.tmpdir(), 'vba-binary-copy-'));
try {
  const node = probe(new NodeFileSystem(), root, 'Node');
  console.log([...memory, ...node].join('\n'));
} finally {
  nodeFs.rmSync(root, { recursive: true, force: true });
}
